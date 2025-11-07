using Helpers;
using Newtonsoft.Json.Linq;
using PDMDA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using static Business_Manager.Controllers.ProductionHistoryController;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Diagnostics;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace Business_Manager.Controllers
{
    public class DrawingController : WebController
    {
        private DrawingDA drawingDA = new DrawingDA();

        //private const string CadFolderPath = @"~/CadFolderPath";
        string CadFolderPath = @"";

        public void GetCadFloderPath()
        {
            int CompanyId = Convert.ToInt32(System.Web.HttpContext.Current.Session["UserCompany"]);

            if (System.Web.HttpContext.Current.Session["CompanySwitch"] != null)
            {
                if (System.Web.HttpContext.Current.Session["CompanySwitch"].ToString() == "manual")
                {
                    CompanyId = Convert.ToInt32(System.Web.HttpContext.Current.Session["CompanyId"]);
                }
            }

            switch (CompanyId)
            {
                case 2:
                    CadFolderPath = @"~/CadFolderPath";
                    break;
                case 4:
                    CadFolderPath = @"~/CadFolderPathDGJMO";
                    break;
                case 3:
                    CadFolderPath = @"~/CadFolderPath";
                    break;
                default:
                    throw new SystemException("此公司別尚未維護設計圖上傳路徑!!");
            }
        }

        #region //View
        public ActionResult CustomerCad()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult RdDesign()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetCustomerCad 取得客戶設計圖資料
        [HttpPost]
        public void GetCustomerCad(int CadId = -1, string CustomerMtlItemNo = "", string CustomerDwgNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerCad", "read,constrained-data");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetCustomerCad(CadId, CustomerMtlItemNo, CustomerDwgNo
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

        #region //GetCustomerCadControl 取得客戶設計圖版本控制資料
        [HttpPost]
        public void GetCustomerCadControl(int ControlId = -1, int CadId = -1, string Edition = ""
            , string StartDate = "", string EndDate = "", string CustomerMtlItemNo = "", string ReleasedStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerCad", "version,constrained-data");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetCustomerCadControl(ControlId, CadId, Edition, StartDate, EndDate, CustomerMtlItemNo, ReleasedStatus
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

        #region //GetCustomerCadEdition 取得客戶設計圖版次資料
        [HttpPost]
        public void GetCustomerCadEdition(int CadId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerCad", "version,constrained-data");

                #region //Request 
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetCustomerCadEdition(CadId);
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

        #region //GetRdDesign 取得研發設計圖資料
        [HttpPost]
        public void GetRdDesign(int DesignId = -1, string CustomerMtlItemNo = "",string MtlItemNo="", int MtlItemId=-1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "read,constrained-data");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetRdDesign(DesignId, CustomerMtlItemNo,MtlItemNo, MtlItemId
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

        #region //GetAiRdDesign 取得 AI  研發設計圖資料 --- MarkChen, 2023-08-24, 以 /Drawing/RdDesign 為模板
        [HttpPost]
        public void GetAiRdDesign(int DesignId = -1, string CustomerMtlItemNo = "", string MtlItemNo = "", string MtlItemName = "", int MtlItemId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "read,constrained-data");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetAiRdDesign(DesignId, CustomerMtlItemNo, MtlItemNo, MtlItemName, MtlItemId
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

        #region //GetRdDesignControl 取得研發設計圖版本控制資料
        [HttpPost]
        public void GetRdDesignControl(int ControlId = -1, int DesignId = -1, string Edition = "", string StartDate = "", string EndDate = "", int MtlItemId = -1, string ReleasedStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "version,constrained-data");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetRdDesignControl(ControlId, DesignId, Edition, StartDate, EndDate, MtlItemId, ReleasedStatus
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

        #region //GetRdDesignControlActiveOnly 取得研發設計圖版本控制資料, 只限 發行狀態=是, by MarkChen 2023.08.24
        [HttpPost]
        public void GetRdDesignControlActiveOnly(int CompanyId=-1,int ControlId = -1, int DesignId = -1, string Edition = "", string StartDate = "", string EndDate = "", int MtlItemId = -1, string ReleasedStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "version,constrained-data");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetRdDesignControlActiveOnly(CompanyId, ControlId, DesignId, Edition, StartDate, EndDate, MtlItemId, ReleasedStatus
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

        #region //GetRdDesignVersion -- 取得研發設計圖系統版次資料
        [HttpPost]
        public void GetRdDesignVersion(int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "version,constrained-data");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetRdDesignVersion(MtlItemId);
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

        #region //GetRdDesignAttribute 取得研發設計圖加工屬性
        [HttpPost]
        public void GetRdDesignAttribute(int ControlId = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "read");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetRdDesignAttribute(ControlId);
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

        #region //GetRdAttributeItem 取得加工屬性資料
        [HttpPost]
        public void GetRdAttributeItem()
        {
            try
            {
                WebApiLoginCheck("RdDesign", "read");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetRdAttributeItem();
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

        #region //GetFolder 取得資料夾項目
        [HttpPost]
        public void GetFolder(string FolderPath = "", int UserId = -1, int ListId = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "read");

                #region //確認此路徑為合法路徑
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetCheckRdWhitelist(ListId, FolderPath, UserId);

                var dataRequestJson = JObject.Parse(dataRequest);
                foreach (var item in dataRequestJson)
                {
                    if (item.Key == "status")
                    {
                        if (item.Value.ToString() == "success")
                        {
                            break;
                        }
                    }
                    else if (item.Key == "msg")
                    {
                        throw new SystemException(item.Value.ToString());
                    }
                }
                #endregion

                string[] Directories = Directory.GetDirectories(FolderPath);

                List<Folder> folders = new List<Folder>();

                foreach (var d in Directories)
                {
                    var dir = new DirectoryInfo(d);
                    var dirName = dir.Name;
                    var dirPath = dir.ToString();
                    var folderInfo = new Folder
                    {
                        FolderName = dir.Name,
                        FolderPath = dir.ToString()
                    };
                    folders.Add(folderInfo);
                }

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    result = folders
                });
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

        #region //GetFiles 搜尋資料夾內檔案
        [HttpPost]
        public void GetFiles(string FolderPath = "", int ListId = -1, int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "read");
                //FolderPath = Path.Combine(ServerFolderPath, FolderPath);
                string[] Files = Directory.GetFiles(FolderPath);

                #region //確認白名單資料是否正確
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetRdWhitelist(ListId, -1, UserId, "");

                jsonResponse = BaseHelper.DAResponse(dataRequest);

                string WhitelistPath = "";
                if (jsonResponse["status"].ToString() == "success")
                {
                    WhitelistPath = jsonResponse["result"][0]["FolderPath"].ToString();
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                List<FileInfoModal> filePaths = new List<FileInfoModal>();
                foreach (var files in Files)
                {
                    var file = new FileInfo(files);
                    string FilePath = file.ToString();

                    #region //處理URL特殊符號
                    if (FilePath.IndexOf("+") != -1) FilePath = FilePath.Replace("+", "%2B");
                    if (FilePath.IndexOf("/") != -1) FilePath = FilePath.Replace("/", "%2F");
                    if (FilePath.IndexOf("?") != -1) FilePath = FilePath.Replace("?", "%3F");
                    if (FilePath.IndexOf("#") != -1) FilePath = FilePath.Replace("#", "%23");
                    if (FilePath.IndexOf("&") != -1) FilePath = FilePath.Replace("&", "%26");
                    if (FilePath.IndexOf("=") != -1) FilePath = FilePath.Replace("=", "%3D");
                    #endregion

                    var fileInfo = new FileInfoModal
                    {
                        FileName = file.Name,
                        FilePath = FilePath,
                        FolderPath = FolderPath,
                        WhitelistPath = WhitelistPath
                    };

                    filePaths.Add(fileInfo);
                }

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    result = filePaths,
                });
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

        #region //GetFileInfo 取得檔案相關資訊
        [HttpPost]
        public void GetFileInfo(string FilePath = "", string MESFolderPath = "", string DowloadFlag = "")
        {
            try
            {
                WebApiLoginCheck("RdDesign", "read,constrained-data");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetFileInfo(FilePath, MESFolderPath, DowloadFlag);
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

        #region //GetRdWhitelist 取得RD白名單資料夾
        [HttpPost]
        public void GetRdWhitelist(int ListId = -1, int DepartmentId = -1, int UserId = -1, string FolderPath = "")
        {
            try
            {
                WebApiLoginCheck("RdDesign", "read,constrained-data");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetRdWhitelist(ListId, DepartmentId, UserId, FolderPath);
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

        #region //GetCheckRdWhitelist 確認路徑是否為合法路徑
        [HttpPost]
        public void GetCheckRdWhitelist(int ListId = -1, string FolderPath = "", int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "read,constrained-data");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetCheckRdWhitelist(ListId, FolderPath, UserId);
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

        #region //GetFileSearch 搜尋指定目錄內所有檔案
        [HttpPost]
        public void GetFileSearch(string FolderPath = "", int ListId = -1, int UserId = -1, string FileName = "")
        {
            try
            {
                WebApiLoginCheck("RdDesign", "read");

                #region //確認此路徑為合法路徑
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetCheckRdWhitelist(ListId, FolderPath, UserId);

                var dataRequestJson = JObject.Parse(dataRequest);
                foreach (var item in dataRequestJson)
                {
                    if (item.Key == "status")
                    {
                        if (item.Value.ToString() == "success")
                        {
                            break;
                        }
                    }
                    else if (item.Key == "msg")
                    {
                        throw new SystemException(item.Value.ToString());
                    }
                }
                #endregion

                #region //確認白名單資料是否正確
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetRdWhitelist(ListId, -1, UserId, "");

                jsonResponse = BaseHelper.DAResponse(dataRequest);

                string WhitelistPath = "";
                if (jsonResponse["status"].ToString() == "success")
                {
                    WhitelistPath = jsonResponse["result"][0]["FolderPath"].ToString();
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                List<FileInfoModal> filePaths = new List<FileInfoModal>();
                List<Folder> folderList = new List<Folder>();

                SearchFilesAndFolders(FolderPath, 0, 0);

                void SearchFilesAndFolders(string searchFolderPath, int searchDepth, int limitCount)
                {
                    if (limitCount >= 100) throw new SystemException("搜尋範圍過大，請縮小搜尋根目錄!!");
                    if (searchDepth >= 10) return; //設定遞迴深度上限

                    #region //先搜尋目前目錄下所有資料夾
                    List<string> Directories = Directory.GetDirectories(searchFolderPath).ToList();
                    if (limitCount == 0) Directories.Add(FolderPath);
                    #endregion

                    foreach (var dir in Directories)
                    {
                        #region //先搜尋當前目錄下是否有符合條件的檔案
                        string[] FilesInfo = Directory.GetFiles(dir);
                        foreach (var fileInfo in FilesInfo)
                        {
                            var file = new FileInfo(fileInfo);
                            if (file.Name.IndexOf(FileName) != -1)
                            {
                                #region //將檔案資訊存入列表
                                string filePath = file.ToString();
                                #region //處理URL特殊符號
                                if (filePath.IndexOf("+") != -1) filePath = filePath.Replace("+", "%2B");
                                if (filePath.IndexOf("/") != -1) filePath = filePath.Replace("/", "%2F");
                                if (filePath.IndexOf("?") != -1) filePath = filePath.Replace("?", "%3F");
                                if (filePath.IndexOf("#") != -1) filePath = filePath.Replace("#", "%23");
                                if (filePath.IndexOf("&") != -1) filePath = filePath.Replace("&", "%26");
                                if (filePath.IndexOf("=") != -1) filePath = filePath.Replace("=", "%3D");
                                #endregion

                                #region //確認此檔案是否已經存在
                                bool checkFileFlag = false;
                                foreach (var item in filePaths)
                                {
                                    if (item.FilePath == filePath)
                                    {
                                        checkFileFlag = true;
                                        break;
                                    }
                                }
                                #endregion

                                if (checkFileFlag == false)
                                {
                                    var fileModal = new FileInfoModal
                                    {
                                        FileName = file.Name,
                                        FilePath = filePath,
                                        FolderPath = FolderPath,
                                        WhitelistPath = WhitelistPath
                                    };

                                    filePaths.Add(fileModal);
                                }
                                #endregion
                            }
                        }
                        #endregion

                        #region //遞迴搜尋
                        limitCount++;
                        SearchFilesAndFolders(dir, searchDepth+1, limitCount);
                        #endregion
                    }
                }

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    result = filePaths,
                });
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

        #region //GetAutoReadCadDesign 取得自動讀CAD圖規格
        [HttpPost]
        public void GetAutoReadCadDesign(int MtlItemId = -1, int CadControlId = 0, int QcItemId = -1, string BallMark = "", string Design = "", string Usl = "", string Lsl = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "read,constrained-data");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetAutoReadCadDesign(MtlItemId, CadControlId, QcItemId, BallMark, Design, Usl, Lsl, OrderBy, PageIndex, PageSize);
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
        #region //AddCustomerCad 新增客戶設計圖資料
        [HttpPost]
        public void AddCustomerCad(string CustomerMtlItemNo = "", string CustomerDwgNo = "")
        {
            try
            {
                WebApiLoginCheck("CustomerCad", "add");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.AddCustomerCad(CustomerMtlItemNo, CustomerDwgNo);
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

        #region //AddCustomerCadControl 新增客戶設計圖版本資料
        [HttpPost]
        public void AddCustomerCadControl(int CadId = -1, string Edition = "", string EditionType = "", string Cause = ""
            ,string ReleasedStatus = "", int CadFile = -1, string OtherFile = "")
        {
            try
            {
                WebApiLoginCheck("CustomerCad", "add");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.AddCustomerCadControl(CadId, Edition, EditionType, Cause, ReleasedStatus, CadFile, OtherFile);
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

        #region //AddCustomerCadControl02 新增客戶設計圖版本02(圖檔上傳方式更改)
        [HttpPost]
        public void AddCustomerCadControl02(int CadId = -1, string Edition = "", string EditionType = "", string Cause = ""
            , string ReleasedStatus = "", string CadFilePath = "", string OtherFile = "")
        {
            try
            {
                WebApiLoginCheck("CustomerCad", "add");

                #region //Request
                drawingDA = new DrawingDA();
                GetCadFloderPath();
                string ServerPath = Server.MapPath(CadFolderPath);
                dataRequest = drawingDA.AddCustomerCadControl02(CadId, Edition, EditionType, Cause, ReleasedStatus, CadFilePath, OtherFile, ServerPath);
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

        #region //AddRdDesign 新增研發設計圖資料
        [HttpPost]
        public void AddRdDesign(string CustomerMtlItemNo = "", int CustomerCadControlId = -1, int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "add");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.AddRdDesign(CustomerMtlItemNo, CustomerCadControlId, MtlItemId);
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

        #region //AddRdDesignControl 新增研發設計圖版本資料
        [HttpPost]
        public void AddRdDesignControl(int DesignId = -1, string Edition = "", string EditionType = "", string Cause = ""
            , string DesignDate = "", string ReleasedStatus = ""
            , int Cad3DFile = -1, int Cad2DFile = -1, int Pdf2DFile = -1, int JmoFile = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "add");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.AddRdDesignControl(DesignId, Edition, EditionType, Cause, DesignDate, ReleasedStatus
                    , Cad3DFile, Cad2DFile, Pdf2DFile, JmoFile);
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

        #region //AddRdDesignControl02 新增研發設計圖版本資料02(圖檔上傳方式更改)
        [HttpPost]
        public void AddRdDesignControl02(int DesignId = -1, string Edition = "", string EditionType = "", string Cause = ""
            , string DesignDate = "", string ReleasedStatus = ""
            , string Cad3DFileAbsolutePath = "", string Cad2DFileAbsolutePath = "", string Pdf2DFileAbsolutePath = "", string JmoFileAbsolutePath = "")
        {
            try
            {
                WebApiLoginCheck("RdDesign", "add");

                #region //Request
                drawingDA = new DrawingDA();
                GetCadFloderPath();
                string ServerPath = Server.MapPath(CadFolderPath);
                dataRequest = drawingDA.AddRdDesignControl02(DesignId, Edition, EditionType, Cause, DesignDate, ReleasedStatus
                    , Cad3DFileAbsolutePath, Cad2DFileAbsolutePath, Pdf2DFileAbsolutePath, JmoFileAbsolutePath, ServerPath);
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

        #region //AddAutoReadCadDesign 新增自動讀CAD圖規格
        [HttpPost]
        public void AddAutoReadCadDesign(int CompanyId = -1, int MtlItemId = -1)
        {
            string Result = "";
            try
            {
                WebApiLoginCheck("RdDesign", "python");
                #region//呼叫Python
                ////取出 dxf_path, product_name, product_ver
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetRdDesignControlActiveOnly(CompanyId, -1, -1, "", "", "", MtlItemId, "", "", 1, 10);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                if (jsonResponse["result"].ToList().Count != 1) throw new SystemException("該品號有超過1個以上的可發行的設計圖，請向RD確認");
                
                if (jsonResponse["status"].ToString() == "success")
                {
                    foreach (var item in jsonResponse["result"])
                    {
                        string Cad2DFileAbsolutePath = item["Cad2DFileAbsolutePath"].ToString();
                        string Edition = item["Edition"].ToString();
                        string CustomerMtlItemNo = item["CustomerMtlItemNo"].ToString();
                        string VersionNo = item["Version"].ToString();
                        //string Cad2DFileAbsolutePath = @"\\192.168.20.199\mes_data\39126A-904-R1&R2.dxf";
                        //string Edition = @"A-01";
                        //string CustomerMtlItemNo = @"39126A-904-R1";

                        //檢查此版本的圖面是否有解析過
                        drawingDA = new DrawingDA();
                        dataRequest = drawingDA.GetAutoReadCadDesignIsDid(CompanyId, MtlItemId, VersionNo);

                        if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                        {
                            string AiCADresult = AiCAD_Python(Cad2DFileAbsolutePath, CustomerMtlItemNo, Edition);//執行AI解析
                            JObject AiCADresultJson = JObject.Parse(AiCADresult);
                            var data = AiCADresultJson["data"] as JObject;
                            if (AiCADresultJson["status"].ToString() == "success")
                            {
                                foreach (var item2 in data.Properties())
                                {
                                    var record = item2.Value as JObject;
                                    string Design = record["Design"].ToString();
                                    string USL = record["USL"].ToString();
                                    string LSL = record["LSL"].ToString();
                                    string QcItemNo = "";
                                    string QcItemName = "";
                                    string QcItemDesc = "";
                                    string BallMark = "";
                                    string Unit = "";
                                    string Remark = "";

                                    if (Design != "" || USL != "" || LSL != "")
                                    {
                                        Result += "/" + Design + ":" + USL + ":" + LSL;
                                        #region //Request
                                        drawingDA = new DrawingDA();
                                        dataRequest = drawingDA.AddAutoReadCadDesign(MtlItemId, -1, Edition, -1, QcItemNo, QcItemName, QcItemDesc, Design, USL, LSL, BallMark, Unit, Remark);
                                        #endregion
                                    }
                                }
                            }
                        }
                        else
                        {
                            throw new SystemException("此品號綁定的圖號版本已有解析規格！");
                        }
                    }
                }
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
            //return Result;
        }
        #endregion

        #region //AddDesign 新增客戶設計圖資料
        [HttpPost]
        public void AddDesign(int MtlItemId = -1, int ControlId = -1, int QcItemId = -1, string QcItemNo = "", string QcItemName = "", string QcItemDesc = "", string DesignValue = "", string UpperTolerance = "", string LowerTolerance = ""
            , string BallMark = "", string Unit = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("RdDesign", "add");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.AddAutoReadCadDesign(MtlItemId, ControlId, "", QcItemId, QcItemNo, QcItemName, QcItemDesc, DesignValue, UpperTolerance, LowerTolerance, BallMark, Unit, Remark);
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
        #region //UpdateCustomerCad 客戶設計圖資料更新
        [HttpPost]
        public void UpdateCustomerCad(int CadId = -1, string CustomerMtlItemNo = "", string CustomerDwgNo = "")
        {
            try
            {
                WebApiLoginCheck("CustomerCad", "update");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.UpdateCustomerCad(CadId, CustomerMtlItemNo, CustomerDwgNo);
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

        #region //UpdateCustomerCadControlReleasedStatus 更新客戶設計圖發行狀態
        [HttpPost]
        public void UpdateCustomerCadControlReleasedStatus(int ControlId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerCad", "released-status");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.UpdateCustomerCadControlReleasedStatus(ControlId);
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

        #region //UpdateRdDesign 研發設計圖資料更新
        [HttpPost]
        public void UpdateRdDesign(int DesignId = -1, string CustomerMtlItemNo = "", int CustomerCadControlId = -1, int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "update");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.UpdateRdDesign(DesignId, CustomerMtlItemNo, CustomerCadControlId, MtlItemId);
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

        #region //UpdateRdDesignControlReleasedStatus 更新研發設計圖發行狀態
        [HttpPost]
        public void UpdateRdDesignControlReleasedStatus(int ControlId = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "released-status");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.UpdateRdDesignControlReleasedStatus(ControlId);
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
        #region //DeleteCustomerCad -- 刪除客戶設計圖資料
        [HttpPost]
        public void DeleteCustomerCad(int CadId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerCad", "delete");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.DeleteCustomerCad(CadId);
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

        #region //DeleteRdDesign -- 刪除研發設計圖資料
        [HttpPost]
        public void DeleteRdDesign(int DesignId = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "delete");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.DeleteRdDesign(DesignId);
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

        #region //Export Excel
        #region //ExportDrawing
        public void ExportDrawing(string Year = "", string Month = "")
        {
            try
            {
                WebApiLoginCheck("RdDesign", "export");

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetDrawingInfo(Year, Month);
                #endregion

                JObject dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() == "success")
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
                    //headerStyle.Fill.BackgroundColor = XLColor.AppleGreen;
                    headerStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    headerStyle.NumberFormat.Format = "@";
                    #endregion

                    #region //dateStyle
                    var dateStyle = XLWorkbook.DefaultStyle;
                    dateStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
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
                    string excelFileName = Year + "年" + Month + "月出圖資料";
                    string excelsheetName = Year + "年" + Month + "月出圖資料";
                    string monthString = Month + "月";
                    string[] headerValues = new string[] { monthString, "客戶設計圖", "研發3D圖", "研發2D圖", "研發JMO" };
                    string colIndexValue = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 15;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 4)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 8).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < headerValues.Length; i++)
                        {
                            colIndexValue = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndexValue).Style.NumberFormat.Format = "@";
                            worksheet.Cell(colIndexValue).Value = headerValues[i];
                            worksheet.Cell(colIndexValue).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, 3)).Value = "次數";
                        worksheet.Cell(BaseHelper.MergeNumberToChar(2, 3)).Value = dataRequestJson["customerCadCount"].ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(3, 3)).Value = dataRequestJson["cad3DFileCount"].ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(4, 3)).Value = dataRequestJson["cad2DFileCount"].ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(5, 3)).Value = dataRequestJson["jmoFileCount"].ToString();
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //置中
                        worksheet.Columns("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        worksheet.Columns("E").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        worksheet.Columns("F").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        //worksheet.Columns("A:C").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Column("4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        #endregion

                        #region //保護及鎖定
                        //workbook.Protect("123");

                        //worksheet.Columns("A:E").Style.Protection.SetLocked(true);
                        worksheet.Columns(BaseHelper.NumberToChar(1) + ":" + BaseHelper.NumberToChar(5)).Style.Protection.SetLocked(true);
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
                        string conditionalFormatRange = BaseHelper.NumberToChar(7) + ":" + BaseHelper.NumberToChar(7);
                        foreach (var range in worksheet.Ranges(conditionalFormatRange))
                        {
                            range.AddConditionalFormat().WhenEquals("M")
                                //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                                .Font.SetFontColor(XLColor.BabyBlue);

                            range.AddConditionalFormat().WhenEquals("F")
                                //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                                .Font.SetFontColor(XLColor.Red);
                        }

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
        #endregion

        #region //API
        #region //AutoDownloadCustomerCad 自動下載客戶設計圖二進制檔案存為實體檔案
        [HttpPost]
        public void AutoDownloadCustomerCad(int CompanyId = -1)
        {
            try
            {
                #region //Request
                drawingDA = new DrawingDA();
                string ServerPath = Server.MapPath(CadFolderPath);
                dataRequest = drawingDA.AutoDownloadCustomerCad(ServerPath, CompanyId);
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

        #region //AutoDownloadRdDesignCad 自動下載RD研發設計圖二進制檔案存為實體檔案
        [HttpPost]
        public void AutoDownloadRdDesignCad(int CompanyId = -1)
        {
            try
            {
                #region //Request
                drawingDA = new DrawingDA();
                string ServerPath = Server.MapPath(CadFolderPath);
                dataRequest = drawingDA.AutoDownloadRdDesignCad(ServerPath, CompanyId);
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

        #region //GetCadFile 取得檔案(絕對路徑方式)
        public virtual ActionResult GetCadFile(string FilePath = "", string DownloadFlag = "")
        {
            try
            {
                if (BaseHelper.ClientLinkType()== "廠內連線") {
                    #region //Request
                    GetCadFloderPath();
                    string ServerPath = Server.MapPath(CadFolderPath);
                    dataRequest = drawingDA.GetFileInfo(FilePath, ServerPath, DownloadFlag);
                    #endregion

                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion

                    if (jsonResponse["status"].ToString() == "success")
                    {
                        string fileName = jsonResponse["result"][0]["FileName"].ToString();
                        byte[] fileContent = (byte[])jsonResponse["result"][0]["FileByte"];
                        string fileExtension = jsonResponse["result"][0]["FileExtension"].ToString();

                        return File(fileContent, FileHelper.GetMime(fileExtension), fileName + fileExtension);
                    }
                    else
                    {
                        throw new SystemException(jsonResponse["msg"].ToString());
                    }
                }
                else {
                    throw new SystemException("此IP非廠內路徑，不可以下載");
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
                Response.ContentType = "application/json";
                Response.Write(jsonResponse.ToString(Newtonsoft.Json.Formatting.None));
                return new EmptyResult();
            }
        }
        #endregion

        #region //AnalyzePdfToText 解析PDF文件轉為字串
        [HttpPost]
        public void AnalyzePdfToText(string FilePath = "")
        {
            try
            {
                string text = "";

                using (PdfReader reader = new PdfReader(FilePath))
                {
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        text += PdfTextExtractor.GetTextFromPage(reader, i);
                    }
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = text
                });
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

        #region //取得檔案
        public virtual ActionResult GetPdmFile(int FileId)
        {
            try
            {
                if (BaseHelper.ClientLinkType() == "廠內連線")
                {
                    return RedirectToAction("GetFile", "Web", new { fileId = FileId });
                }
                else
                {
                    throw new SystemException("此IP非廠內路徑，不可以下載");
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

                Response.ContentType = "application/json";
                Response.Write(jsonResponse.ToString(Newtonsoft.Json.Formatting.None));

                return new EmptyResult();
            }
        }
        #endregion

        #region//取得製令設計圖資訊
        [HttpPost]
        [Route("api/Drawing/DrawingData")]     
        public void GetDrawingData(string CompanyNo = "", string SecretKey = "", int MoId = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "GetDrawingData");
                #endregion

                #region //Request
                drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetDrawingData(MoId);
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

        #region//人工智慧解析CAD產出尺寸
        public string AiCAD_Python(string dxf_path,string product_name,string product_ver)
        {
            string result = "", error = "";
            string programPath = @"C:\Python\Python311\AI_label_MES\main_recog_Lable-v3.1.3.py";

            try
            {
                ProcessStartInfo processStartInfo = new ProcessStartInfo
                {
                    Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\"", programPath, dxf_path, product_name, product_ver),
                    CreateNoWindow = true,
                    FileName = @"C:\Python\Python311\python.exe",
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    UseShellExecute = false
                };
                using (Process process = Process.Start(processStartInfo))
                {
                    using (StreamReader readerOutput = process.StandardOutput)
                    {
                        result = readerOutput.ReadToEnd().Replace("\r", "").Replace("\n", "");
                    }
                    if (result.Equals(""))
                    {
                        using (StreamReader readerError = process.StandardError)
                        {
                            error = readerError.ReadToEnd();
                            logger.Error("status:error:" + error );
                        }
                    }
                    if (error.Length > 0) result += error;
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
                result = jsonResponse.ToString();
                logger.Error(e.Message);
            }
            return result;
        }
        #endregion

        #endregion
    }
}