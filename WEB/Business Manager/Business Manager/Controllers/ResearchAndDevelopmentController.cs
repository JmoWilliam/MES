using ClosedXML.Excel;
using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PDMDA;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class ResearchAndDevelopmentController : WebController
    {
        private ResearchAndDevelopmentDA researchAndDevelopmentDA = new ResearchAndDevelopmentDA();

        public ActionResult DesignPatentManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult LensFOVManagement()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult ModelDesignManagement()
        {
            ViewLoginCheck();
            return View();
        }

        #region//Get
        #region //GetUserAuthority 取得專利使用者權限
        [HttpPost]
        public void GetUserAuthority(string UserNo = "")
        {
            try
            {
                WebApiLoginCheck("DesignPatentManagement", "read");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.GetUserAuthority(UserNo);
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

        #region //GetDesignPatentExcel 取得專利Exce資料
        [HttpPost]
        public void GetDesignPatentExcel(
             int DpId = -1, string ApplicationNo = "", string ApplicationDate = "", string Applicant = "",
             string PublicationNo = "", string PublicationDate = "", string AnnouncementNo = "", string AnnouncementDate = "",
             string ImplementationExample = "", string LensQuantity = "", string AperturePosition = "", string HalfFieldofView = "", string ImageHeight = "",
             string L1RefractivePower = "", string L1aSurfaceType = "", string L1bSurfaceType = "", string L2RefractivePower = "", string L2aSurfaceType = "", string L2bSurfaceType = "",
             string L3RefractivePower = "", string L3aSurfaceType = "", string L3bSurfaceType = "", string L4RefractivePower = "", string L4aSurfaceType = "", string L4bSurfaceType = "",
             string L5RefractivePower = "", string L5aSurfaceType = "", string L5bSurfaceType = "", string L6RefractivePower = "", string L6aSurfaceType = "", string L6bSurfaceType = "",
             string L7RefractivePower = "", string L7aSurfaceType = "", string L7bSurfaceType = "", string L8RefractivePower = "", string L8aSurfaceType = "", string L8bSurfaceType = "",
             string L9RefractivePower = "", string L9aSurfaceType = "", string L9bSurfaceType = "", string LXRefractivePower = "", string LXaSurfaceType = "", string LXbSurfaceType = "",
             string OrderBy ="", int PageIndex = -1, int PageSize=-1)
        {
            try
            {
                WebApiLoginCheck("DesignPatentManagement", "read");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.GetDesignPatentExcel(
                    DpId, ApplicationNo, ApplicationDate, Applicant, 
                    PublicationNo, PublicationDate, AnnouncementNo, AnnouncementDate, 
                    ImplementationExample, LensQuantity, AperturePosition, HalfFieldofView, ImageHeight, 
                    L1RefractivePower, L1aSurfaceType, L1bSurfaceType, L2RefractivePower, L2aSurfaceType, L2bSurfaceType, 
                    L3RefractivePower, L3aSurfaceType, L3bSurfaceType, L4RefractivePower, L4aSurfaceType, L4bSurfaceType, 
                    L5RefractivePower, L5aSurfaceType, L5bSurfaceType, L6RefractivePower, L6aSurfaceType, L6bSurfaceType,
                    L7RefractivePower, L7aSurfaceType, L7bSurfaceType, L8RefractivePower, L8aSurfaceType, L8bSurfaceType,
                    L9RefractivePower, L9aSurfaceType, L9bSurfaceType, LXRefractivePower, LXaSurfaceType, LXbSurfaceType,
                    OrderBy, PageIndex, PageSize);
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

        #region //GetDesignPatentCount 取得專利Exce資料總比數
        [HttpPost]
        public void GetDesignPatentCount()
        {
            try
            {
                WebApiLoginCheck("DesignPatentManagement", "read");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.GetDesignPatentCount();
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

        #region //GetModelDesign 取得鏡頭設計資料
        [HttpPost]
        public void GetModelDesign(int MdId = -1, string Model = "", string MdIdList = "",
              string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ModelDesignManagement", "read");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.GetModelDesign(MdId, Model, MdIdList,
                    OrderBy, PageIndex, PageSize);
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

        #region //GetLensFOV 取得鏡頭FOV資料
        [HttpPost]
        public void GetLensFOV(
             string Model = "", int LfId = -1, string ModelName = "", double FOV = -1, double RealHeight = -1,
             double GreaterMeets = -1, double HFOV = -1, double VFOV = -1, double DFOV = -1,string ButtonType = "",
             double FOVRangeD = -1, double FOVRangeT = -1,　string ModelNoList = "",
             string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("LensFOVManagement", "read");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.GetLensFOV(
                    Model,LfId, ModelName, FOV, RealHeight,
                    GreaterMeets, HFOV, VFOV, DFOV, ButtonType,
                    FOVRangeD, FOVRangeT, ModelNoList,
                    OrderBy, PageIndex, PageSize);
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
        #region //AddDesignPatentExcel 上傳專利Exce資料
        [HttpPost]
        public void AddDesignPatentExcel(string DpIdList = "")
        {
            try
            {
                WebApiLoginCheck("DesignPatentManagement", "add");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.AddDesignPatentExcel(DpIdList);
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

        #region //AddModelDesign 新增鏡頭設計資料
        [HttpPost]
        public void AddModelDesign(
            string Model = "", string Construction = "", double EFL = -1, double FNo = -1
            , double TTL = -1, double CRA = -1, string OpticalDistortionFTan = "", string OpticalDistortionF = ""
            , double MaxImageCircle = -1, string Sensor = "", double VFOV = -1, double HFOV = -1, double DFOV = -1
            , string PartName = "", string IR = "", string MechanicalRetainer = "", double MechanicalFBL = -1, string MechanicalThread = ""
            , string IPRating = "", string Stage = ""
            )
        {
            try
            {
                WebApiLoginCheck("ModelDesignManagement", "add");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.AddModelDesign(
                     Model, Construction, EFL, FNo, TTL, CRA, OpticalDistortionFTan,  
                     OpticalDistortionF, MaxImageCircle, Sensor, VFOV, HFOV, DFOV,
                     PartName, IR, MechanicalRetainer, MechanicalFBL, MechanicalThread,
                     IPRating, Stage
                    );
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

        #region //AddModelDesignExcel 上傳鏡頭設計Exce資料
        [HttpPost]
        public void AddModelDesignExcel(string ModelDesignList = "")
        {
            try
            {
                WebApiLoginCheck("ModelDesignManagement", "excel-add");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.AddModelDesignExcel(ModelDesignList);
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

        #region //AddLensFOV 新增LensFOV資料
        [HttpPost]
        public void AddLensFOV(string Model = "", double FOV = -1, double Realheight = -1)
        {
            try
            {
                WebApiLoginCheck("LensFOVManagement", "add");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.AddLensFOV(Model, FOV, Realheight);
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

        #region //AddLensFOVExcel 上傳LensFOVExce資料
        [HttpPost]
        public void AddLensFOVExcel(string LensFOVList = "")
        {
            try
            {
                WebApiLoginCheck("LensFOVManagement", "excel-add");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.AddLensFOVExcel(LensFOVList);
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
        #region //UpdateDesignPatentExcel 更新專利Exce資料
        [HttpPost]
        public void UpdateDesignPatentExcel(string DpIdList = "")
        {
            try
            {
                WebApiLoginCheck("DesignPatentManagement", "update");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.UpdateDesignPatentExcel(DpIdList);
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

        #region //UpdateModelDesign 更新機種設計資料
        [HttpPost]
        public void UpdateModelDesign(int MdId = -1,
            string Model = "", string Construction = "", double EFL = -1, double FNo = -1
            , double TTL = -1, double CRA = -1, string OpticalDistortionFTan = "", string OpticalDistortionF = ""
            , double MaxImageCircle = -1, string Sensor = "", double VFOV = -1, double HFOV = -1, double DFOV = -1
            , string PartName = "", string IR = "", string MechanicalRetainer = "", double MechanicalFBL = -1, string MechanicalThread = ""
            , string IPRating = "", string Stage = ""
            )
        {
            try
            {
                WebApiLoginCheck("ModelDesignManagement", "update");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.UpdateModelDesign(MdId,
                     Model, Construction, EFL, FNo, TTL, CRA, OpticalDistortionFTan,
                     OpticalDistortionF, MaxImageCircle, Sensor, VFOV, HFOV, DFOV,
                     PartName, IR, MechanicalRetainer, MechanicalFBL, MechanicalThread,
                     IPRating, Stage
                    );
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

        #region //UpdateLensFOV 更新專利Exce資料
        [HttpPost]
        public void UpdateLensFOV(
             int LfId = -1, string Model = "", double FOV = -1, double Realheight = -1)
        {
            try
            {
                WebApiLoginCheck("LensFOVManagement", "update");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.UpdateLensFOV(LfId, Model, FOV, Realheight);
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
        #region //DelteDesignPatentExcel 刪除專利Exce資料
        [HttpPost]
        public void DelteDesignPatentExcel(string DpIdList = "")
        {
            try
            {
                WebApiLoginCheck("DesignPatentManagement", "delete");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.DelteDesignPatentExcel(DpIdList);
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

        #region //DeleteModelDesign 刪除機種設計資料
        [HttpPost]
        public void DeleteModelDesign(int MdId = -1)
        {
            try
            {
                WebApiLoginCheck("ModelDesignManagement", "delete");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.DeleteModelDesign(MdId);
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

        #region //DeleteLensFOV 刪除鏡頭FOV資料
        [HttpPost]
        public void DeleteLensFOV(int LfId = -1)
        {
            try
            {
                WebApiLoginCheck("LensFOVManagement", "delete");

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.DeleteLensFOV(LfId);
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

        #region//PDF
        #endregion

        #region //ExcelModelDesign FOV機種列表匯出Excel 
        public void ExcelModelDesign(string Model = "", int LfId = -1, string ModelName = "", double FOV = -1, double RealHeight = -1
            , double GreaterMeets = -1, double HFOV = -1, double VFOV = -1, double DFOV = -1,string ButtonType = ""
            , double FOVRangeD = -1, double FOVRangeT = -1 , int MdId = -1, string ModelNoList = "", string Model2 = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("LensFOVManagement", "excel");
                List<string> LensFOVList = new List<string>();
                
                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.GetLensFOV(
                    Model, LfId, ModelName, FOV, RealHeight,
                    GreaterMeets, HFOV, VFOV, DFOV, ButtonType,
                    FOVRangeD, FOVRangeT, ModelNoList,
                    OrderBy, PageIndex, PageSize);
                dynamic[] dataFOVId = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                if (dataFOVId.Count() <= 0) throw new SystemException("查無單據!無法執行匯出");
                foreach (var item in dataFOVId)
                {
                    LensFOVList.Add(item.Model.ToString());
                }
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    #region //樣式

                    #region //defaultStyle
                    var defaultStyle = XLWorkbook.DefaultStyle;
                    defaultStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    defaultStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    defaultStyle.Border.TopBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.RightBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.TopBorderColor = XLColor.NoColor;
                    defaultStyle.Border.BottomBorderColor = XLColor.NoColor;
                    defaultStyle.Border.LeftBorderColor = XLColor.NoColor;
                    defaultStyle.Border.RightBorderColor = XLColor.NoColor;
                    defaultStyle.Fill.BackgroundColor = XLColor.NoColor;
                    defaultStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    defaultStyle.Font.FontSize = 12;
                    defaultStyle.Font.Bold = false;
                    defaultStyle.Protection.SetLocked(false);
                    #endregion

                    #region //titleStyle
                    var titleStyle = XLWorkbook.DefaultStyle;
                    titleStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    titleStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    titleStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.TopBorderColor = XLColor.Black;
                    titleStyle.Border.BottomBorderColor = XLColor.Black;
                    titleStyle.Border.LeftBorderColor = XLColor.Black;
                    titleStyle.Border.RightBorderColor = XLColor.Black;
                    titleStyle.Fill.BackgroundColor = XLColor.NoColor;
                    titleStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    titleStyle.Font.FontSize = 16;
                    titleStyle.Font.Bold = false;
                    #endregion

                    #region //headerStyle
                    var headerStyle = XLWorkbook.DefaultStyle;
                    headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    headerStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.TopBorderColor = XLColor.Black;
                    headerStyle.Border.BottomBorderColor = XLColor.Black;
                    headerStyle.Border.LeftBorderColor = XLColor.Black;
                    headerStyle.Border.RightBorderColor = XLColor.Black;
                    headerStyle.Fill.BackgroundColor = XLColor.AliceBlue;
                    headerStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dateStyle
                    var dateStyle = XLWorkbook.DefaultStyle;
                    dateStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    dateStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dateStyle.Border.TopBorder = XLBorderStyleValues.None;
                    dateStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    dateStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    dateStyle.Border.RightBorder = XLBorderStyleValues.None;
                    dateStyle.Border.TopBorderColor = XLColor.NoColor;
                    dateStyle.Border.BottomBorderColor = XLColor.NoColor;
                    dateStyle.Border.LeftBorderColor = XLColor.NoColor;
                    dateStyle.Border.RightBorderColor = XLColor.NoColor;
                    dateStyle.Fill.BackgroundColor = XLColor.NoColor;
                    dateStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    dateStyle.Font.FontSize = 12;
                    dateStyle.Font.Bold = false;
                    dateStyle.DateFormat.Format = "yyyy-mm-dd HH:mm:ss";
                    #endregion

                    #region //numberStyle
                    var numberStyle = XLWorkbook.DefaultStyle;
                    numberStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    numberStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    numberStyle.Border.TopBorder = XLBorderStyleValues.None;
                    numberStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    numberStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    numberStyle.Border.RightBorder = XLBorderStyleValues.None;
                    numberStyle.Border.TopBorderColor = XLColor.NoColor;
                    numberStyle.Border.BottomBorderColor = XLColor.NoColor;
                    numberStyle.Border.LeftBorderColor = XLColor.NoColor;
                    numberStyle.Border.RightBorderColor = XLColor.NoColor;
                    numberStyle.Fill.BackgroundColor = XLColor.NoColor;
                    numberStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    numberStyle.Font.FontSize = 12;
                    numberStyle.Font.Bold = false;
                    numberStyle.NumberFormat.Format = "_-* #,##0.00_-;-* #,##0.00_-;_-* \" - \"??_-;_-@_-";
                    #endregion

                    #region //currencyStyle
                    var currencyStyle = XLWorkbook.DefaultStyle;
                    currencyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    currencyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    currencyStyle.Border.TopBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.RightBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.TopBorderColor = XLColor.NoColor;
                    currencyStyle.Border.BottomBorderColor = XLColor.NoColor;
                    currencyStyle.Border.LeftBorderColor = XLColor.NoColor;
                    currencyStyle.Border.RightBorderColor = XLColor.NoColor;
                    currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                    currencyStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    currencyStyle.Font.FontSize = 12;
                    currencyStyle.Font.Bold = false;
                    currencyStyle.DateFormat.Format = "$#,##0.00";
                    #endregion

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】鏡頭FOV機種列表匯出Excel";
                    string excelsheetName = "鏡頭FOV機種列表匯出Excel1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "機種名稱", "H-FOV", "V-FOV", "D-FOV", "Construction", "EFL(mm)", "F/No", "TTL", "CRA", "OpticalDistortion", "OpticalDistortion", "MaxImageCircle", "Sensor", "VFOV", "HFOV", "DFOV", "PartName", "IR", "Mechanical-Retainer", "Mechanical-F.B.L.", "Mechanical-Thread", "IPRating", "Stage"};
                    string colIndex = "";
                    int rowIndex = 1;
                    switch (ButtonType)
                    {
                        case "AllH":
                        case "HFOV":
                            int indexToRemove = Array.FindIndex(header, item => item == "D-FOV");
                            if (indexToRemove != -1)
                            {
                                string[] newHeader = new string[header.Length - 1];
                                Array.Copy(header, 0, newHeader, 0, indexToRemove);
                                Array.Copy(header, indexToRemove + 1, newHeader, indexToRemove, header.Length - indexToRemove - 1);
                                header = newHeader;
                            }
                            break;
                    }
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        foreach (var item in data)
                        {
                            rowIndex++;

                            if (rowIndex-1 <= dataFOVId.Count())
                            {
                                int DFOVRemove = 0;
                                worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.Model.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.HFOVIH.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.VFOVIH.ToString();
                                switch (ButtonType)
                                {
                                    case "AllH":
                                    case "HFOV":
                                        DFOVRemove = 1;
                                        break;
                                    case "AllD":
                                    case "DFOV":
                                        worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.DFOVIH.ToString();
                                        break;
                                }
                                worksheet.Cell(BaseHelper.MergeNumberToChar(5-DFOVRemove, rowIndex)).Value = item.Construction.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(6-DFOVRemove, rowIndex)).Value = item.EFL.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(7-DFOVRemove, rowIndex)).Value = item.FNo.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(8-DFOVRemove, rowIndex)).Value = item.TTL.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(9- DFOVRemove, rowIndex)).Value = item.CRA.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(10-DFOVRemove, rowIndex)).Value = item.OpticalDistortionFTan.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(11-DFOVRemove, rowIndex)).Value = item.OpticalDistortionF.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(12-DFOVRemove, rowIndex)).Value = item.MaxImageCircle.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(13-DFOVRemove, rowIndex)).Value = item.Sensor.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(14-DFOVRemove, rowIndex)).Value = item.VFOV.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(15-DFOVRemove, rowIndex)).Value = item.HFOV.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(16-DFOVRemove, rowIndex)).Value = item.DFOV.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(17-DFOVRemove, rowIndex)).Value = item.PartName.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(18-DFOVRemove, rowIndex)).Value = item.IR.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(19-DFOVRemove, rowIndex)).Value = item.MechanicalRetainer.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(20-DFOVRemove, rowIndex)).Value = item.MechanicalFBL.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(21-DFOVRemove, rowIndex)).Value = item.MechanicalThread.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(22-DFOVRemove, rowIndex)).Value = item.IPRating.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(23- DFOVRemove, rowIndex)).Value = item.Stage.ToString();
                                
                            }
                        }

                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 14;
                        //worksheet.Column(2).Width = 16;

                        #endregion

                        #region //置中
                        //worksheet.Columns("A:C").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Columns("A:C").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Column("4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        #endregion

                        #region //保護及鎖定
                        //workbook.Protect("123");
                        //worksheet.Protect("123");

                        //worksheet.Columns("A:E").Style.Protection.SetLocked(true);
                        //worksheet.Columns(BaseHelper.NumberToChar(1) + ":" + BaseHelper.NumberToChar(5)).Style.Protection.SetLocked(true);
                        //worksheet.Cell(4, 1).Style.Protection.SetLocked(true);
                        //worksheet.Range(3, 1, 8, 1).Style.Protection.SetLocked(true);
                        //worksheet.Range(BaseHelper.NumberToChar(1) + ":" + BaseHelper.NumberToChar(5)).Style.Protection.SetLocked(true);
                        #endregion

                        #region //合併儲存格
                        //worksheet.Range(8, 1, 8, 9).Merge();
                        #endregion

                        #region //公式
                        ////R=Row、C=Column
                        ////[]指相對位置，[-]向前，[+]向後
                        ////未[]指絕對位置
                        //worksheet.Cell(10, 3).FormulaR1C1 = "RC[-2]+R3C2";
                        //worksheet.Cell(11, 3).FormulaR1C1 = "RC[-2]+RC[2]";
                        //worksheet.Cell(12, 3).FormulaR1C1 = "RC[-2]+RC2";
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(1);
                        #endregion

                        #region //隱藏
                        //worksheet.Column(1).Hide();
                        #endregion

                        #region //格式化
                        //string conditionalFormatRange = "G:G";
                        //string conditionalFormatRange = BaseHelper.NumberToChar(7) + ":" + BaseHelper.NumberToChar(7);
                        //foreach (var range in worksheet.Ranges(conditionalFormatRange))
                        //{
                        //    range.AddConditionalFormat().WhenEquals("M")
                        //        //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                        //        .Font.SetFontColor(XLColor.BabyBlue);

                        //    range.AddConditionalFormat().WhenEquals("F")
                        //        //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                        //        .Font.SetFontColor(XLColor.Red);
                        //}

                        //string conditionalFormatRange2 = "E:E";
                        //foreach (var range in worksheet.Ranges(conditionalFormatRange2))
                        //{
                        //    range.AddConditionalFormat().WhenContains("1")
                        //        //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                        //        .Font.SetFontColor(XLColor.RedNcs);
                        //}
                        #endregion

                        #region //複製
                        //worksheet.CopyTo("複製");
                        #endregion

                        #endregion

                        #region //表格
                        //var fakeData = Enumerable.Range(1, 5)
                        //.Select(x => new FakeData
                        //{
                        //    Time = TimeSpan.FromSeconds(x * 123.667),
                        //    X = x,
                        //    Y = -x,
                        //    Address = "a" + x,
                        //    Distance = x * 100
                        //}).ToArray();

                        //var table = worksheet.Cell(10, 1).InsertTable(fakeData);
                        //table.Style.Font.FontSize = 9;
                        //var tableData = worksheet.Cell(17, 1).InsertData(fakeData);
                        //tableData.Style.Font.FontSize = 9;
                        //worksheet.Range(11, 1, 21, 1).Style.DateFormat.Format = "HH:mm:ss.000";
                        #endregion

                        #region //EXCEL匯出
                        using (MemoryStream output = new MemoryStream())
                        {
                            workbook.SaveAs(output);

                            excelFile = output.ToArray();
                        }
                        #endregion
                    }

                    string fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = excelFile;


                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        fileGuid,
                        fileName = excelFileName,
                        fileExtension = ".xlsx"
                    });
                    #endregion


                }
                else
                {
                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion
                }
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

        #region //ExcelLensFOV 鏡頭FOV資料列表匯出Excel 
        public void ExcelLensFOV(
             string Model = "", int LfId = -1, string ModelName = "", double FOV = -1, double RealHeight = -1
            , double GreaterMeets = -1, double HFOV = -1, double VFOV = -1, double DFOV = -1, string ButtonType = ""
            , double FOVRangeD = -1, double FOVRangeT = -1, string ModelNoList = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("LensFOVManagement", "excel");
                List<string> LensFOVList = new List<string>();

                #region //Request
                researchAndDevelopmentDA = new ResearchAndDevelopmentDA();
                dataRequest = researchAndDevelopmentDA.GetLensFOV(
                    Model, LfId, ModelName, FOV, RealHeight,
                    GreaterMeets, HFOV, VFOV, DFOV, ButtonType,
                    FOVRangeD, FOVRangeT, ModelNoList,
                    OrderBy, PageIndex, PageSize);
                #endregion

                

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    #region //樣式

                    #region //defaultStyle
                    var defaultStyle = XLWorkbook.DefaultStyle;
                    defaultStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    defaultStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    defaultStyle.Border.TopBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.RightBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.TopBorderColor = XLColor.NoColor;
                    defaultStyle.Border.BottomBorderColor = XLColor.NoColor;
                    defaultStyle.Border.LeftBorderColor = XLColor.NoColor;
                    defaultStyle.Border.RightBorderColor = XLColor.NoColor;
                    defaultStyle.Fill.BackgroundColor = XLColor.NoColor;
                    defaultStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    defaultStyle.Font.FontSize = 12;
                    defaultStyle.Font.Bold = false;
                    defaultStyle.Protection.SetLocked(false);
                    #endregion

                    #region //titleStyle
                    var titleStyle = XLWorkbook.DefaultStyle;
                    titleStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    titleStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    titleStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.TopBorderColor = XLColor.Black;
                    titleStyle.Border.BottomBorderColor = XLColor.Black;
                    titleStyle.Border.LeftBorderColor = XLColor.Black;
                    titleStyle.Border.RightBorderColor = XLColor.Black;
                    titleStyle.Fill.BackgroundColor = XLColor.NoColor;
                    titleStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    titleStyle.Font.FontSize = 16;
                    titleStyle.Font.Bold = false;
                    #endregion

                    #region //headerStyle
                    var headerStyle = XLWorkbook.DefaultStyle;
                    headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    headerStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.TopBorderColor = XLColor.Black;
                    headerStyle.Border.BottomBorderColor = XLColor.Black;
                    headerStyle.Border.LeftBorderColor = XLColor.Black;
                    headerStyle.Border.RightBorderColor = XLColor.Black;
                    headerStyle.Fill.BackgroundColor = XLColor.AliceBlue;
                    headerStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dateStyle
                    var dateStyle = XLWorkbook.DefaultStyle;
                    dateStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    dateStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dateStyle.Border.TopBorder = XLBorderStyleValues.None;
                    dateStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    dateStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    dateStyle.Border.RightBorder = XLBorderStyleValues.None;
                    dateStyle.Border.TopBorderColor = XLColor.NoColor;
                    dateStyle.Border.BottomBorderColor = XLColor.NoColor;
                    dateStyle.Border.LeftBorderColor = XLColor.NoColor;
                    dateStyle.Border.RightBorderColor = XLColor.NoColor;
                    dateStyle.Fill.BackgroundColor = XLColor.NoColor;
                    dateStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    dateStyle.Font.FontSize = 12;
                    dateStyle.Font.Bold = false;
                    dateStyle.DateFormat.Format = "yyyy-mm-dd HH:mm:ss";
                    #endregion

                    #region //numberStyle
                    var numberStyle = XLWorkbook.DefaultStyle;
                    numberStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    numberStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    numberStyle.Border.TopBorder = XLBorderStyleValues.None;
                    numberStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    numberStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    numberStyle.Border.RightBorder = XLBorderStyleValues.None;
                    numberStyle.Border.TopBorderColor = XLColor.NoColor;
                    numberStyle.Border.BottomBorderColor = XLColor.NoColor;
                    numberStyle.Border.LeftBorderColor = XLColor.NoColor;
                    numberStyle.Border.RightBorderColor = XLColor.NoColor;
                    numberStyle.Fill.BackgroundColor = XLColor.NoColor;
                    numberStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    numberStyle.Font.FontSize = 12;
                    numberStyle.Font.Bold = false;
                    numberStyle.NumberFormat.Format = "_-* #,##0.00_-;-* #,##0.00_-;_-* \" - \"??_-;_-@_-";
                    #endregion

                    #region //currencyStyle
                    var currencyStyle = XLWorkbook.DefaultStyle;
                    currencyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    currencyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    currencyStyle.Border.TopBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.RightBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.TopBorderColor = XLColor.NoColor;
                    currencyStyle.Border.BottomBorderColor = XLColor.NoColor;
                    currencyStyle.Border.LeftBorderColor = XLColor.NoColor;
                    currencyStyle.Border.RightBorderColor = XLColor.NoColor;
                    currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                    currencyStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    currencyStyle.Font.FontSize = 12;
                    currencyStyle.Font.Bold = false;
                    currencyStyle.DateFormat.Format = "$#,##0.00";
                    #endregion

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】鏡頭FOV資料彙整檔";
                    string excelsheetName = "鏡頭FOV資料彙整頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "機種名稱", "H-FOV", "V-FOV", "D-FOV"};
                    string colIndex = "";
                    int rowIndex = 1;
                    switch (ButtonType)
                    {
                        case "AllH":
                        case "HFOV":
                            Array.Resize(ref header, header.Length - 1);
                            break;
                    }
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        //rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        foreach (var item in data)
                        {
                            rowIndex++;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.Model.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.HFOVIH.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.VFOVIH.ToString();
                            switch (ButtonType)
                            {
                                case "AllD":
                                case "DFOV":
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.DFOVIH.ToString();
                                    break;
                            }
                        }

                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 14;
                        //worksheet.Column(2).Width = 16;

                        #endregion

                        #region //置中
                        //worksheet.Columns("A:C").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Columns("A:C").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Column("4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        #endregion

                        #region //保護及鎖定
                        //workbook.Protect("123");
                        //worksheet.Protect("123");

                        //worksheet.Columns("A:E").Style.Protection.SetLocked(true);
                        //worksheet.Columns(BaseHelper.NumberToChar(1) + ":" + BaseHelper.NumberToChar(5)).Style.Protection.SetLocked(true);
                        //worksheet.Cell(4, 1).Style.Protection.SetLocked(true);
                        //worksheet.Range(3, 1, 8, 1).Style.Protection.SetLocked(true);
                        //worksheet.Range(BaseHelper.NumberToChar(1) + ":" + BaseHelper.NumberToChar(5)).Style.Protection.SetLocked(true);
                        #endregion

                        #region //合併儲存格
                        //worksheet.Range(8, 1, 8, 9).Merge();
                        #endregion

                        #region //公式
                        ////R=Row、C=Column
                        ////[]指相對位置，[-]向前，[+]向後
                        ////未[]指絕對位置
                        //worksheet.Cell(10, 3).FormulaR1C1 = "RC[-2]+R3C2";
                        //worksheet.Cell(11, 3).FormulaR1C1 = "RC[-2]+RC[2]";
                        //worksheet.Cell(12, 3).FormulaR1C1 = "RC[-2]+RC2";
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(1);
                        #endregion

                        #region //隱藏
                        //worksheet.Column(1).Hide();
                        #endregion

                        #region //格式化
                        //string conditionalFormatRange = "G:G";
                        //string conditionalFormatRange = BaseHelper.NumberToChar(7) + ":" + BaseHelper.NumberToChar(7);
                        //foreach (var range in worksheet.Ranges(conditionalFormatRange))
                        //{
                        //    range.AddConditionalFormat().WhenEquals("M")
                        //        //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                        //        .Font.SetFontColor(XLColor.BabyBlue);

                        //    range.AddConditionalFormat().WhenEquals("F")
                        //        //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                        //        .Font.SetFontColor(XLColor.Red);
                        //}

                        //string conditionalFormatRange2 = "E:E";
                        //foreach (var range in worksheet.Ranges(conditionalFormatRange2))
                        //{
                        //    range.AddConditionalFormat().WhenContains("1")
                        //        //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                        //        .Font.SetFontColor(XLColor.RedNcs);
                        //}
                        #endregion

                        #region //複製
                        //worksheet.CopyTo("複製");
                        #endregion

                        #endregion

                        #region //表格
                        //var fakeData = Enumerable.Range(1, 5)
                        //.Select(x => new FakeData
                        //{
                        //    Time = TimeSpan.FromSeconds(x * 123.667),
                        //    X = x,
                        //    Y = -x,
                        //    Address = "a" + x,
                        //    Distance = x * 100
                        //}).ToArray();

                        //var table = worksheet.Cell(10, 1).InsertTable(fakeData);
                        //table.Style.Font.FontSize = 9;
                        //var tableData = worksheet.Cell(17, 1).InsertData(fakeData);
                        //tableData.Style.Font.FontSize = 9;
                        //worksheet.Range(11, 1, 21, 1).Style.DateFormat.Format = "HH:mm:ss.000";
                        #endregion

                        #region //EXCEL匯出
                        using (MemoryStream output = new MemoryStream())
                        {
                            workbook.SaveAs(output);

                            excelFile = output.ToArray();
                        }
                        #endregion
                    }

                    string fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = excelFile;


                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        fileGuid,
                        fileName = excelFileName,
                        fileExtension = ".xlsx"
                    });
                    #endregion

                    #region //刪除製令條碼QR圖片
                    //DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/MoId/QR"));

                    //FileInfo[] files = di.GetFiles();
                    //foreach (FileInfo file in files)
                    //{
                    //    file.Delete();
                    //}
                    #endregion

                    //#region //刪除製令條碼1D圖片
                    //DirectoryInfo di2 = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/MoId/1D"));

                    //FileInfo[] files2 = di2.GetFiles();
                    //foreach (FileInfo file in files2)
                    //{
                    //    file.Delete();
                    //}
                    //#endregion

                }
                else
                {
                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion
                }
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