using BPMDA;
using Helpers;
using IOTDA;
using MESDA;
using SCMDA;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Org.BouncyCastle.Asn1.Ocsp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Web.Mvc;
using ZXing;
using ZXing.Common;

namespace Business_Manager.Controllers
{
    public class IOSController : WebController
    {
        private MesBasicInformationDA mesBasicInformationDA = new MesBasicInformationDA();
        private DeptDocumentDA deptDocumentDA = new DeptDocumentDA();
        private MachineAvailabilityDA machineAvailabilityDA = new MachineAvailabilityDA();
        private SaleOrderDA saleOrderDA = new SaleOrderDA();
        //public static LinePushNotifiHelper linePushNotifiHelper = new Helpers.LinePushNotifiHelper();


        public ActionResult Index()
        {
            return View();
        }

        #region //Api
        #region //AddDocumentSynchronize -- 文件簽核資料建立
        [HttpPost]
        [Route("api/BPM/AddDeptDocument")]
        public void AddDocumentSynchronize(string Company = "", string SecretKey = "",
            string FolderNo = "", string SenderUserNo = "", string DocName = "", int SenderId = -1, string SendTime = "",
            string ApproverUserNo = "", int ApproverId = -1, string ApproverTime = "", string Remark = "", string Status = "N")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "AddDocumentSynchronize");
                #endregion

                #region //Request
                dataRequest = deptDocumentDA.AddDocumentSynchronize(Company, FolderNo, SenderUserNo, DocName,
                    SenderId, SendTime, ApproverUserNo, ApproverId, ApproverTime, Remark, Status);
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

        #region //UpdateDocumentSynchronize -- 文件簽核資料同步
        [HttpPost]
        [Route("api/BPM/UpdateDeptDocument")]
        public void UpdateDocumentSynchronize(string Company = "", string SecretKey = "",
            string UserNo = "", string FolderNo = "", string DocName = "", int SenderId = -1, string SendTime = "",
            string ApproverUserNo = "", int ApproverId = -1, string ApproverTime = "", string Remark = "", string Status = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateDocumentSynchronize");
                #endregion

                #region //Request
                dataRequest = deptDocumentDA.UpdateDocumentSynchronize(Company, UserNo, FolderNo, DocName,
                    SenderId, SendTime, ApproverUserNo, ApproverId, ApproverTime, Remark, Status);
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

        #region //InquiryDocumentSynchronize -- 文件簽核資料查詢
        [HttpPost]
        [Route("api/BPM/InquiryDeptDocument")]
        public void InquiryDocumentSynchronize(string Company = "", string SecretKey = "",
            string FolderNo = "", string SenderUserNo = "", string DocName = "", int SenderId = -1, string SendTime = "",
            string ApproverUserNo = "", int ApproverId = -1, string ApproverTime = "", string Remark = "", string Status = "N")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "InquiryDocumentSynchronize");
                #endregion

                #region //Request
                dataRequest = deptDocumentDA.InquiryDocumentSynchronize(Company, FolderNo, SenderUserNo, DocName,
                    SenderId, SendTime, ApproverUserNo, ApproverId, ApproverTime, Remark, Status);
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

        #region //InquiryDocumentSynchronize -- 推播訊息(尚未完成編寫)
        [HttpPost]
        [Route("api/BPM/DemandMessageDeptDocument")]
        public void DemandMessage(string Company = "", string SecretKey = "", string FolderNo = "",
            string txType = "", int MessageType = -1, string DemandMessageContent = "", string SendUserNo = "", string SendStatus = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "DemandMessage");
                #endregion

                #region //Request
                dataRequest = deptDocumentDA.DemandMessage(Company, FolderNo, txType, MessageType, DemandMessageContent, SendUserNo, SendStatus);
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


        #region //APP用
        #region //--mes製令在製查詢   
        [HttpPost]
        [Route("api/MFG/GetManufacturingDetilIniOS")]
        public void GetManufacturingDetilIniOS(string Company = "",int CompanyId = -1, string SecretKey = "", int ModeId = -1, string WoErpNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = "")
        {
            try
            {
                ApiKeyVerify(Company, SecretKey, "GetManufacturingDetilIniOS");

                dataRequest = deptDocumentDA.GetMESProcessForiOS(CompanyId, ModeId, WoErpNo, MtlItemNo,
                   StartDate, EndDate);

                jsonResponse = BaseHelper.DAResponse(dataRequest);

            }

            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //--取得管理職名單
        [HttpPost]
        [Route("api/BPA/GetDepartmentLeader")]
        public void GetDepartmentLeader(string Company = "", string SecretKey = "")
        {
            try
            {
                ApiKeyVerify(Company, SecretKey, "GetDepartmentLeader");

                dataRequest = deptDocumentDA.GetDepartmentLeader();

                jsonResponse = BaseHelper.DAResponse(dataRequest);

            }

            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //--iOS登入
        [HttpPost]
        [Route("api/Login/LoginiOS")]
        public void LoginiOS(string Account, string Password)
        {
            try
            {
                //string captchaCode = Session["CaptchaCode"].ToString();
                //if (!VerifyCode.Equals(captchaCode)) throw new SystemException("驗證碼錯誤!");

                #region //Request
                dataRequest = basicInformationDA.GetLogin(Account, Password);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    string passwordStatus = result[0]["PasswordStatus"].ToString();
                    string userId = result[0]["UserId"].ToString();

                    if (passwordStatus == "N")
                    {
                        if (Convert.ToDateTime(result[0]["PasswordExpire"]) > DateTime.Now)
                        {
                            

                            LoginLog(Convert.ToInt32(result[0]["UserId"]), Account);

                            #region //Response
                            jsonResponse = JObject.FromObject(new{status = "success",msg = "登入成功!",result = userId});
                            //jsonResponse = BaseHelper.DAResponse(dataRequest);
                            #endregion
                        }
                        else
                        {
                            string iv = BaseHelper.RandomCode(16);
                            string key = BaseHelper.StrMid(BaseHelper.Sha256Encrypt(Account), 12, 32);

                            Session["NewLoginIV"] = iv;
                            Session["NewLogin"] = BaseHelper.AESEncrypt(Account, key, iv);
                            Session.Timeout = 300;

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "expirePassword",
                               // userId = Convert.ToInt32(result[0]["UserId"]),
                                msg = "密碼過期",
                                result = "密碼過期"
                            });
                            #endregion
                        }
                    }
                    else
                    {
                        string iv = BaseHelper.RandomCode(16);
                        string key = BaseHelper.StrMid(BaseHelper.Sha256Encrypt(Account), 12, 32);

                        Session["NewLoginIV"] = iv;
                        Session["NewLogin"] = BaseHelper.AESEncrypt(Account, key, iv);
                        Session.Timeout = 300;

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "newPassword",
                           // userId = Convert.ToInt32(result[0]["UserId"]),
                            msg = "首次登入",
                            result = "首次登入"
                        });
                        #endregion
                    }
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

        #region //iOS新密碼登入
        [HttpPost]
        [Route("api/Login/NewLoginiOS")]
        public void NewLoginiOS(string Account, string NewPassword, string ConfirmPassword)
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.UpdateNewLogin(Account, NewPassword, ConfirmPassword);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    LoginLog(Convert.ToInt32(result[0]["UserId"]), Account);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "登入成功!",
                        result = "登入成功"
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

        #region //UpdatePassword 密碼修改
        [HttpPost]
        [Route("api/Login/UpdatePasswordiOS")]
        public void UpdatePasswordiOS(int UserId = -1, string NewPassword = "", string ConfirmPassword = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                if (UserId <= 0) UserId = Convert.ToInt32(Session["UserId"]);

                dataRequest = basicInformationDA.UpdatePassword(UserId, NewPassword, ConfirmPassword);
                #endregion

                #region //Response
                //jsonResponse = BaseHelper.DAResponse(dataRequest);
                jsonResponse = JObject.FromObject(new
                {
                    status = "PasswordChange",
                    msg = "修改完成",
                    result = "密碼修改完成，請重新登入"
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

        #region //GetCustomer 取得客戶資料
        [HttpPost]
        [Route("api/BPA/GetCustomeriOS")]
        public void GetCustomeriOS(string Company = "", string SecretKey = "",int CustomerId = -1, string CustomerNo = "", string CustomerName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("CustomerManagement", "read, constrained-data");
                ApiKeyVerify(Company, SecretKey, "GetCustomeriOS");

                #region //Request
                dataRequest = deptDocumentDA.GetCustomeriOS(CustomerId, CustomerNo, CustomerName, Status
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

       
        #region //--取得功能
        [HttpPost]
        [Route("api/BPA/GetAPPFunction")]
        public void GetAPPFunction(string Company = "", string SecretKey = "", int UserId = -1)
        {
            try
            {
                ApiKeyVerify(Company, SecretKey, "GetAPPFunction");

                dataRequest = deptDocumentDA.GetAPPFunction(UserId);

                jsonResponse = BaseHelper.DAResponse(dataRequest);

            }

            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //--出貨計畫查詢
        [HttpPost]
        [Route("api/BPA/GetDeliveriOS")]
        public void GetDeliveriOS(string Company = "", int CompanyId = -1, string SecretKey = "", string DeliveryStatus = "", string SoIds = "", string SoErpFullNo = "", string CustomerMtlItemNo = "", int CustomerId = -1
            , string MtlItemNo = "", int SalesmenId = -1, string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify(Company, SecretKey, "GetDeliveriOS");

                #region //Request
                dataRequest = deptDocumentDA.GetDeliveriOS(Company, CompanyId, DeliveryStatus, SoIds, SoErpFullNo, CustomerMtlItemNo, CustomerId, MtlItemNo, SalesmenId, StartDate, EndDate
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

        #region //GetSoErpFullNo 取得訂單單別+單號
        [HttpPost]
        [Route("api/SCM/GetSoErpFullNoiOS")]
        public void GetSoErpFullNoiOS(string Company = "", string SecretKey = "", int SoId = -1, string SoErpPrefix = "", string SoErpNo = "", string SoErpFullNo = "", string DeliveryStatus = ""
            , string OrderBy = "")
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");
                ApiKeyVerify(Company, SecretKey, "GetSoErpFullNoiOS");
                #region //Request
                dataRequest = saleOrderDA.GetSoErpFullNo(DeliveryStatus);
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

        #region //歷史紀錄稼動率(趨勢折線圖)
        [HttpPost]
        [Route("api/IOT/getFactoryDayHistory")]
        public void GetFactoryDayHistory(string Company = "", string SecretKey = "", int ID = -1)
        {
            try
            {
                ApiKeyVerify(Company, SecretKey, "getFactoryDayHistory");

                dataRequest = machineAvailabilityDA.GetFactoryDayHistory(ID);

                jsonResponse = BaseHelper.DAResponse(dataRequest);

            }

            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //機台機時(長條圖)
        [HttpPost]
        [Route("api/IOT/getFactoryTimeHistory")]
        public void GetFactoryTimeHistory(string Company = "", string SecretKey = "", int ID = -1, string StartTime = "", string EndTime = "")
        {
            try
            {
                ApiKeyVerify(Company, SecretKey, "getFactoryTimeHistory");

                dataRequest = machineAvailabilityDA.GetFactoryTimeHistory(ID, StartTime, EndTime);

                jsonResponse = BaseHelper.DAResponse(dataRequest);

            }

            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //機台即時狀態列表
        [HttpPost]
        [Route("api/IOT/GetFactoryRealTimeList")]
        public void GetFactoryRealTimeList(string Company = "", string SecretKey = "")
        {
            try
            {
                ApiKeyVerify(Company, SecretKey, "GetFactoryRealTimeList");

                dataRequest = machineAvailabilityDA.GetFactoryRealTimeList();

                jsonResponse = BaseHelper.DAResponse(dataRequest);

            }

            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProdMode 取得生產模式
        [HttpPost]
        [Route("api/MES/GetProdMode")]
        public void GetProdMode(string Company = "", string SecretKey = "", int ModeId = -1, string ModeNo = "", string ModeName = "", string Status = "", string BarcodeCtrl = "", string ScrapRegister = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ProdModeManagment", "read");
                ApiKeyVerify(Company, SecretKey, "GetProdMode");
                #region //Request
                dataRequest = mesBasicInformationDA.GetProdMode(ModeId, ModeNo, ModeName, Status, BarcodeCtrl, ScrapRegister
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

        #region //GetUser 取得使用者資料
        [HttpPost]
        [Route("api/Login/GetUserInfo")]
        public void GetUserInfo(string Company = "", string SecretKey = "",int UserId = -1, int DepartmentId = -1, int CompanyId = -1, string Departments = ""
            , string UserNo = "", string UserName = "", string Gender = "", string Status = "", string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("UserManagement", "read,constrained-data");
                ApiKeyVerify(Company, SecretKey, "GetUserInfo");
                #region //Request
                dataRequest = basicInformationDA.GetUser(UserId, DepartmentId, CompanyId, Departments
                    , UserNo, UserName, Gender, Status, SearchKey
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

        #region //取得使用者列表
        [HttpPost]
        [Route("api/BPA/GetUserList")]
        public void GetUserList(string Company = "", string SecretKey = "")
        {
            try
            {
                ApiKeyVerify(Company, SecretKey, "GetUserList");

                dataRequest = deptDocumentDA.GetUserList();

                jsonResponse = BaseHelper.DAResponse(dataRequest);

            }

            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //取得公司代碼
        [HttpPost]
        [Route("api/BPA/GetUserCompany")]
        public void GetUserCompany(string Company = "", string SecretKey = "", int UserId = -1)
        {
            try
            {
                ApiKeyVerify(Company, SecretKey, "GetUserCompany");

                dataRequest = deptDocumentDA.GetUserCompany(UserId);

                jsonResponse = BaseHelper.DAResponse(dataRequest);

            }

            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //登出
        [HttpPost]
        [Route("api/Login/LogoutiOS")]
        public void Logout(string UserNo = ""/*, string KeyText = ""*/)
        {
            try
            {
                #region //Request
                dataRequest = deptDocumentDA.DeleteUserLoginKey(UserNo/*, KeyText*/);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                
                #endregion

                Session.Clear();
                //Response.Cookies["Login"].Expires = DateTime.Now.AddDays(-1);
                //Response.Cookies["LoginKey"].Expires = DateTime.Now.AddDays(-1);
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

        #region//GETTOKEN
        [HttpPost]
        [Route("api/Login/GetUserToken")]
        public void GetUserToken(string Company = "", string SecretKey = "", int UserId = -1)
        {
            try
            {
                ApiKeyVerify(Company, SecretKey, "GetUserToken");

                dataRequest = deptDocumentDA.GetUserToken(UserId);

                jsonResponse = BaseHelper.DAResponse(dataRequest);

            }

            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion
        #endregion


        #region //Line推播

        #region //LineLogin 跳轉LINE登入畫面
        [HttpPost]
        [Route("api/LinePush/LineLogin")]
        public void LineLogin()
        {
            try
            {
                //string ClientID = "qJ1mhzvc0cRF1NKKBdLd0G";
                //string callbackURL = "http://192.168.134.52:16668/Event/LineNotifyTest2";
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //Request

                #region //跳轉LINE登入畫面
                dataRequest = deptDocumentDA.GetUserNotifyAuthorize();
                #endregion

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

        #region //GetLineNotifyToken 取Token
        [HttpPost]
        [Route("api/LinePush/GetLineNotifyToken")]
        public void GetLineNotifyToken(string usercode)
        {
            try
            {
                //string ClientID = "qJ1mhzvc0cRF1NKKBdLd0G";
                //string callbackURL = "http://192.168.134.52:16668/Event/LineNotifyTest2";
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "RequestMamoFunction");
                #endregion

                #region //Request

                #region //跳轉LINE登入畫面
                dataRequest = deptDocumentDA.GetLineNotifyToken(usercode);
                #endregion

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


        #endregion
    }
}