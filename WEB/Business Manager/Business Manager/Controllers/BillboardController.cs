using EBPDA;
using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class BillboardController : WebController
    {
        private BillboardDA billboardDA = new BillboardDA();

        #region //View
        public ActionResult Announcement()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetBoard 取得公佈欄資料
        [HttpPost]
        public void GetBoard(int BoardId = -1, int BoardTypeId = -1, string Title = "", string Status = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("Announcement", "read");

                #region //Request
                billboardDA = new BillboardDA();
                dataRequest = billboardDA.GetBoard(BoardId, BoardTypeId, Title, Status, StartDate, EndDate
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

        #region //GetBoardFile 取得公佈欄檔案
        [HttpPost]
        public void GetBoardFile(int BoardId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("Announcement", "read");

                #region //Request
                billboardDA = new BillboardDA();
                dataRequest = billboardDA.GetBoardFile(BoardId
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
        #region //AddBoard 公佈欄資料新增
        [HttpPost]
        public void AddBoard(int BoardTypeId = -1, string Title = "", string AnnounceDate = "", string Content = "")
        {
            try
            {
                WebApiLoginCheck("Announcement", "add");

                #region //Request
                billboardDA = new BillboardDA();
                dataRequest = billboardDA.AddBoard(BoardTypeId, Title, AnnounceDate, Content);
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

        #region //AddBoardFile 公佈欄檔案新增
        [HttpPost]
        public void AddBoardFile(int BoardId = -1, int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("Announcement", "add");

                #region //Request
                billboardDA = new BillboardDA();
                dataRequest = billboardDA.AddBoardFile(BoardId, FileId);
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
        #region //UpdateBoard 公佈欄資料更新
        [HttpPost]
        public void UpdateBoard(int BoardId = -1, int BoardTypeId = -1, string Title = "", string AnnounceDate = "", string Content = "")
        {
            try
            {
                WebApiLoginCheck("Announcement", "update");

                #region //Request
                billboardDA = new BillboardDA();
                dataRequest = billboardDA.UpdateBoard(BoardId, BoardTypeId, Title, AnnounceDate, Content);
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

        #region //UpdateBoardStatus 公佈欄狀態更新
        [HttpPost]
        public void UpdateBoardStatus(int BoardId = -1)
        {
            try
            {
                WebApiLoginCheck("Announcement", "status-switch");

                #region //Request
                billboardDA = new BillboardDA();
                dataRequest = billboardDA.UpdateBoardStatus(BoardId);
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

        #region //UpdateBoardFileType 公佈欄檔案類別更新
        [HttpPost]
        public void UpdateBoardFileType(int BoardFileId = -1, string FileType = "")
        {
            try
            {
                WebApiLoginCheck("Announcement", "update");

                #region //Request
                billboardDA = new BillboardDA();
                dataRequest = billboardDA.UpdateBoardFileType(BoardFileId, FileType);
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
        #region //DeleteBoard 公佈欄資料刪除
        [HttpPost]
        public void DeleteBoard(int BoardId = -1)
        {
            try
            {
                WebApiLoginCheck("Announcement", "delete");

                #region //Request
                billboardDA = new BillboardDA();
                dataRequest = billboardDA.DeleteBoard(BoardId);
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

        #region //DeleteBoardFile 公佈欄檔案刪除
        [HttpPost]
        public void DeleteBoardFile(int BoardFileId = -1)
        {
            try
            {
                WebApiLoginCheck("Announcement", "delete");

                #region //Request
                billboardDA = new BillboardDA();
                dataRequest = billboardDA.DeleteBoardFile(BoardFileId);
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