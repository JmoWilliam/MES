using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using System.Web.Mvc;
using Dapper;
using System.Linq;
using QMSDA;

namespace Business_Manager.Controllers
{
    public class QmsScheduleController : WebController
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private readonly QmsScheduleDA qmsScheduleDA = new QmsScheduleDA();
        

        #region //View
        /// <summary>
        /// 机台操作页面 - 无需系统登录，仅验证员工号
        /// </summary>
        public ActionResult Index(int? id = null, int? companyId = null)
        {
                    if (id.HasValue)
                    {
                ViewBag.MachineId = id.Value;
                    }
            if (companyId.HasValue)
            {
                ViewBag.CompanyId = companyId.Value;
            }
            return View();
        }

        /// <summary>
        /// 数据查询页面 - 需要系统登录
        /// </summary>
        public ActionResult Query()
        {
			// 登录检查：需要存在 Login 和 LoginKey Cookie
			HttpCookie loginCookie = Request.Cookies.Get("Login");
			HttpCookie loginKeyCookie = Request.Cookies.Get("LoginKey");
			if (loginCookie == null || string.IsNullOrEmpty(loginCookie.Value)
				|| loginKeyCookie == null || string.IsNullOrEmpty(loginKeyCookie.Value))
			{
				string returnUrl = Request?.Url != null ? Server.UrlEncode(Request.Url.PathAndQuery) : Server.UrlEncode("/QmsSchedule/Query");
				return Redirect("/User/Login?ReturnUrl=" + returnUrl);
			}

            try
            {
                string userNo = loginCookie.Value;
                var userInfo = qmsScheduleDA.GetUserOrganizationByUserNo(userNo);

                string userName = userInfo?.UserName ?? userNo;
                string departmentName = userInfo?.DepartmentName ?? "";
                string companyName = userInfo?.CompanyName ?? "";

                ViewBag.UserNo = userNo;
                ViewBag.UserName = userName;
                ViewBag.DepartmentName = departmentName;
                ViewBag.CompanyName = companyName;

                string departmentPart = string.IsNullOrWhiteSpace(departmentName) ? "未分配部门" : departmentName;
                string companyPart = string.IsNullOrWhiteSpace(companyName) ? "未分配公司" : companyName;
                ViewBag.UserDisplayName = $"{userName}（{departmentPart} / {companyPart}）";
            }
            catch (Exception ex)
            {
                logger.Error(ex, "加载用户组织信息失败");
                ViewBag.UserDisplayName = loginCookie.Value;
            }

            return View();
        }
        #endregion

        #region //基础数据 API

        #region //ValidateEmployee 验证员工号
        /// <summary>
        /// 验证员工号是否存在
        /// </summary>
        [HttpPost]
        public void ValidateEmployee(string employeeNo = "")
        {
            try
            {
                if (string.IsNullOrEmpty(employeeNo))
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "员工号不能为空"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                var result = qmsScheduleDA.ValidateEmployee(employeeNo);

                if (result != null)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "验证成功",
                        result = result
                    });
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "员工号不存在"
                    });
                }
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

        #region //GetMachineModes 获取机型列表
        /// <summary>
        /// 获取机型列表 (QV-1, QV-2, QV-3, QV-4)
        /// </summary>
        [HttpPost]
        public void GetMachineModes()
        {
            try
            {
                var result = qmsScheduleDA.GetMachineModes().ToList();

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    result = result
                });
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

        #region //GetMachinesByMode 获取指定机型的机台列表
        /// <summary>
        /// 获取指定机型的机台列表
        /// </summary>
        [HttpPost]
        public void GetMachinesByMode(int qcMachineModeId = -1)
        {
            try
            {
                var result = qmsScheduleDA.GetMachinesByMode(qcMachineModeId).ToList();

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        result = result
                    });
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

        #region //GetMachineInfo 获取机台详情
        /// <summary>
        /// 获取机台详情 (用于 URL ?id=MachineId)
        /// </summary>
        [HttpPost]
        public void GetMachineInfo(int machineId = -1)
        {
            try
            {
                if (machineId <= 0)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "机台ID无效"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                var result = qmsScheduleDA.GetMachineInfo(machineId);

                if (result != null)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        result = result
                    });
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "机台不存在"
                    });
                }
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

        #region //量测单据查询 API

        #region //SearchQcRecordById 根据量测单号查询
        /// <summary>
        /// 根据量测单号(QcRecordId)查询量测单据
        /// </summary>
        [HttpPost]
        public void SearchQcRecordById(string qcRecordId = "", int? companyId = null)
        {
            try
            {
                if (string.IsNullOrEmpty(qcRecordId))
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "量测单号不能为空"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                var result = qmsScheduleDA.SearchQcRecordById(qcRecordId, companyId);

                if (result != null)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "查询成功",
                        result = result
                    });
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "未找到该量测单号"
                    });
                }
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

        #region //量测单据 API

        #region //GetQcRecords 获取量测单据列表
        /// <summary>
        /// 获取量测单据列表
        /// </summary>
        [HttpPost]
        public void GetQcRecords(int qcRecordId = -1, string status = "", int pageIndex = -1, int pageSize = -1)
        {
            try
            {
                // 简单的登录检查
                HttpCookie loginCookie = Request.Cookies.Get("Login");
                if (loginCookie == null || string.IsNullOrEmpty(loginCookie.Value))
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "请先登录"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                var result = qmsScheduleDA.GetQcRecords(qcRecordId, status, pageIndex, pageSize).ToList();

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    result = result
                });
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

        #region //GetQcRecordDetail 获取量测单详情
        /// <summary>
        /// 获取单个量测单详情
        /// </summary>
        [HttpPost]
        public void GetQcRecordDetail(int qcRecordId = -1)
        {
            try
            {
                if (qcRecordId <= 0)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "送检单号无效"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                var result = qmsScheduleDA.GetQcRecordDetail(qcRecordId);

                if (result != null)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        result = result
                    });
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "送检单不存在"
                    });
                }
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

        #region //机台排程管理 API

        #region //GetScheduleList 获取机台排程列表
        /// <summary>
        /// 获取指定机台的排程列表
        /// </summary>
        [HttpPost]
        public void GetScheduleList(int machineId = -1, string status = "", int pageIndex = -1, int pageSize = -1)
        {
            try
            {
                var result = qmsScheduleDA.GetScheduleList(machineId, status, pageIndex, pageSize).ToList();

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    data = result
                });
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

        #region //AddToSchedule 收件 - 添加到机台排程
        /// <summary>
        /// 收件 - 将送检单添加到机台排程
        /// </summary>
        [HttpPost]
        public void AddToSchedule(int qcRecordId = -1, int machineId = -1, int userId = -1)
        {
            try
            {
                if (qcRecordId <= 0 || machineId <= 0 || userId <= 0)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "参数无效"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                qmsScheduleDA.AddToSchedule(qcRecordId, machineId, userId);

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "收件成功"
                });
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

        #region //RejectQcRecord 驳回送检单
        /// <summary>
        /// 驳回送检单
        /// </summary>
        [HttpPost]
        public void RejectQcRecord(int qcRecordId = -1, string reason = "", int userId = -1)
        {
            try
            {
                if (qcRecordId <= 0 || userId <= 0)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "参数无效"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                if (string.IsNullOrEmpty(reason))
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "驳回原因不能为空"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                var affected = qmsScheduleDA.RejectQcRecord(qcRecordId, reason, userId);

                if (affected > 0)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "驳回成功"
                    });
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "送检单不存在"
                    });
                }
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

        #region //StartMeasurement 开始量测
        /// <summary>
        /// 开始量测
        /// </summary>
        [HttpPost]
        public void StartMeasurement(int qmmId = -1, int executorId = -1)
        {
            try
            {
                if (qmmId <= 0 || executorId <= 0)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "参数无效"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                qmsScheduleDA.StartMeasurement(qmmId, executorId);

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "开始量测成功"
                });
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

        #region //EndMeasurement 结束量测
        /// <summary>
        /// 结束量测
        /// </summary>
        [HttpPost]
        public void EndMeasurement(int qmmId = -1, int userId = -1)
        {
            try
            {
                if (qmmId <= 0 || userId <= 0)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "参数无效"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                qmsScheduleDA.EndMeasurement(qmmId, userId);

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "结束量测成功"
                });
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

        #region //ChangeMachine 更换机台
        /// <summary>
        /// 更换机台
        /// </summary>
        [HttpPost]
        public void ChangeMachine(int qmmId = -1, int toMachineId = -1, int userId = -1)
        {
            try
            {
                if (qmmId <= 0 || toMachineId <= 0 || userId <= 0)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "参数无效"
                    });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                qmsScheduleDA.ChangeMachine(qmmId, toMachineId, userId);

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "更换机台成功"
                });
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

        #region //机台状态 API

        #region //GetMachineStatus 获取机台状态统计
        /// <summary>
        /// 获取机台状态统计（每个机台的排队数和量测中数量）
        /// </summary>
        [HttpPost]
        public void GetMachineStatus(int qcMachineModeId = -1)
        {
            try
            {
                var result = qmsScheduleDA.GetMachineStatus(qcMachineModeId).ToList();

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    result = result
                });
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

        /// <summary>
        /// 根据送检单号查询机台作业列表
        /// </summary>
        [HttpPost]
        public void QueryOrderByNumber(string orderNumber = "", int page = 1, int pageSize = 10)
        {
            try
            {
                if (string.IsNullOrEmpty(orderNumber))
                {
                    jsonResponse = JObject.FromObject(new { status = "error", msg = "送检单号不能为空" });
                    Response.Write(jsonResponse.ToString());
                    return;
                }

                var data = qmsScheduleDA.QueryOrderByNumber(orderNumber, page, pageSize);

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "查询成功",
                    data = data
                });
            }
            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new { status = "error", msg = e.Message });
                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        /// <summary>
        /// 获取所有机台类型
        /// </summary>
        [HttpPost]
        public void GetAllMachineTypes()
        {
            try
            {
                var result = qmsScheduleDA.GetAllMachineTypes().ToList();

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    result = result
                });
            }
            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new { status = "error", msg = e.Message });
                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        /// <summary>
        /// 根据机台类型获取机台状态
        /// </summary>
        [HttpPost]
        public void GetMachineStatusByType(int qcMachineModeId = 0)
        {
            try
            {
                var result = qmsScheduleDA.GetMachineStatusByType(qcMachineModeId).ToList();

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    result = result
                });
            }
            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new { status = "error", msg = e.Message });
                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #endregion
    }
}

