using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Web;

namespace PDMDA
{
    public class DfmQuotation
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";

        public int CurrentCompany = -1;
        public int CurrentUser = -1;
        public int CreateBy = -1;
        public int LastModifiedBy = -1;
        public DateTime CreateDate = default(DateTime);
        public DateTime LastModifiedDate = default(DateTime);

        public string sql = "";
        public JObject jsonResponse = new JObject();
        public Logger logger = LogManager.GetCurrentClassLogger();
        public DynamicParameters dynamicParameters = new DynamicParameters();
        public SqlQuery sqlQuery = new SqlQuery();

        public DfmQuotation()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];

            GetUserInfo();
            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;

            dynamicParameters = new DynamicParameters();
            sqlQuery = new SqlQuery();
        }

        #region //GetUserInfo 取得使用者資訊
        private void GetUserInfo()
        {
            try
            {
                CurrentCompany = Convert.ToInt32(HttpContext.Current.Session["UserCompany"]);
                CurrentUser = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                CreateBy = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                LastModifiedBy = Convert.ToInt32(HttpContext.Current.Session["UserId"]);

                if (HttpContext.Current.Session["CompanySwitch"] != null)
                {
                    if (HttpContext.Current.Session["CompanySwitch"].ToString() == "manual")
                    {
                        CurrentCompany = Convert.ToInt32(HttpContext.Current.Session["CompanyId"]);
                    }
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }
        }
        #endregion

        #region//Get
        #endregion

        #region//Add
        #endregion

        #region//Update

        #endregion

        #region//Delete
        #endregion

        #region //Transaction
        public string TxQuotationCalculation(int DfmId)
        {
            if (DfmId == -1) throw new Exception("DfmId 不可為空");

            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region 刪除舊有報價計算資料
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.DfmQiSolutionProcess
                                  FROM PDM.DfmQiSolutionProcess a
                                       INNER JOIN PDM.DfmQiProcess b ON a.DfmQiProcessId = b.DfmQiProcessId
	                                   INNER JOIN PDM.DfmQuotationItem c ON b.DfmQiId = c.DfmQiId
	                                   INNER JOIN PDM.DesignForManufacturing d ON c.DfmId = d.DfmId
                                 WHERE d.DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            DfmId
                        });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion



                        #region //判斷DfmId是否存在, 且資料是否已維護完整
                        sql = @"SELECT a.DfmNo,b.DfmQuotationName,b.MfgFlag,b.StandardCost,
                                       b.MaterialStatus,b.DfmQiOSPStatus,b.DfmQiProcessStatus,
                                       a.RfqDetailId
                                  FROM PDM.DesignForManufacturing a
                                       INNER JOIN PDM.DfmQuotationItem b ON a.DfmId = b.DfmId
                                 WHERE a.DfmId = @DfmId
                                   AND (b.MaterialStatus = 'A' or b.DfmQiOSPStatus = 'A' or b.DfmQiProcessStatus = 'A')";
                        dynamicParameters.Add("DfmId", DfmId);
                        var unCompleteCnt = sqlConnection.Query(sql, dynamicParameters);
                        if (unCompleteCnt.Count() > 0) throw new SystemException("DFM還未製程或物料尚未維護完成");
                        #endregion

                        #region //取出報價項目
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.DfmQiId,a.DfmNo,b.DfmQuotationName,b.MfgFlag,b.StandardCost,
                                       b.MaterialStatus,b.DfmQiOSPStatus,b.DfmQiProcessStatus,
                                       a.RfqDetailId
                                  FROM PDM.DesignForManufacturing a
                                       INNER JOIN PDM.DfmQuotationItem b ON a.DfmId = b.DfmId
                                 WHERE a.DfmId = @DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        var QuotationItemResult = sqlConnection.Query(sql, dynamicParameters);
                        int QuotationResult = 0;
                        foreach(var item in QuotationItemResult)
                        {
                            int RfqDetailId = Convert.ToInt32(item.RfqDetailId);
                            string MaterialStatus = item.MaterialStatus;
                            string DfmQiOSPStatus = item.DfmQiOSPStatus;
                            string DfmQiProcessStatus = item.DfmQiProcessStatus;
                            int DfmQiId = Convert.ToInt32(item.DfmQiId);

                            #region //找出報價方案
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.RfqLineSolutionId,a.SortNumber,a.SolutionQty
                                              FROM SCM.RfqLineSolution a
                                             WHERE a.RfqDetailId = @RfqDetailId";
                            dynamicParameters.Add("RfqDetailId", RfqDetailId);
                            var LineSolutionResult = sqlConnection.Query(sql, dynamicParameters);
                            if (LineSolutionResult.Count() <= 0) throw new SystemException("沒有報價方案資料!");
                            foreach (var item3 in LineSolutionResult)
                            {
                                int RfqLineSolutionId = Convert.ToInt32(item3.RfqLineSolutionId);
                                int SolutionSortNumber = Convert.ToInt32(item3.SortNumber);
                                int SolutionQty = Convert.ToInt32(item3.SolutionQty);
                                double ResourceAmt = 0;
                                double OverheadAmt = 0;
                                int DfmQiSolutionId = -1;

                                #region //取得 DfmQiSolutionId
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT DfmQiSolutionId
                                                  FROM PDM.DfmQiSolution
                                                 WHERE DfmQiId = @DfmQiId
                                                   AND RfqLineSolutionId = @RfqLineSolutionId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);
                                dynamicParameters.Add("RfqLineSolutionId", RfqLineSolutionId);
                                var DfmQiSolutionResult = sqlConnection.Query(sql, dynamicParameters);
                                if (DfmQiSolutionResult.Count() > 0)
                                {
                                    DfmQiSolutionId = Convert.ToInt32(sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).DfmQiSolutionId);
                                }
                                else
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO PDM.DfmQiSolution (DfmQiId, RfqLineSolutionId
                                                ,CreateDate, CreateBy
                                                ,LastModifiedDate, LastModifiedBy)
                                                OUTPUT INSERTED.DfmQiSolutionId
                                                VALUES (@DfmQiId, @RfqLineSolutionId
                                                ,@CreateDate, @CreateBy
                                                ,@LastModifiedDate, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DfmQiId,
                                            RfqLineSolutionId,
                                            CreateDate,
                                            CreateBy = CreateBy,
                                            LastModifiedDate,
                                            LastModifiedBy = CreateBy
                                        });
                                    var insertDfmQiSolution = sqlConnection.Query(sql, dynamicParameters);
                                    foreach (var itemQiSolution in insertDfmQiSolution)
                                    {
                                        DfmQiSolutionId = Convert.ToInt32(itemQiSolution.DfmQiSolutionId);
                                    }
                                }

                                #endregion

                                #region //計算物料費用
                                if (MaterialStatus.Equals("Y"))
                                {
                                    QuotationResult++;

                                    #region //更新物料成本
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE PDM.DfmQiSolution
                                               SET MaterialAmt = b.Amount,
                                                   LastModifiedDate = @LastModifiedDate,
                                                   LastModifiedBy = @LastModifiedBy
                                              FROM PDM.DfmQiSolution a
                                                   INNER JOIN (SELECT x.DfmQiId,SUM(x.Amount) Amount
	                                                             FROM PDM.DfmQiMaterial x
		                                                        GROUP BY x.DfmQiId) b ON a.DfmQiId = b.DfmQiId
                                             WHERE a.DfmQiId = @DfmQiId";
                                    dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DfmQiId,
                                        LastModifiedDate = DateTime.Now,
                                        LastModifiedBy = CreateBy
                                    });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion

                                #region //計算外包費用(先不執行)
                                #endregion

                                #region //計算人工與製造費用
                                if (DfmQiProcessStatus.Equals("Y"))
                                {
                                    QuotationResult++;
                                    #region //找出DfmQiProcess
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.DfmQiId,DfmQiProcessId,a.SortNumber,a.ParameterId,a.AllocationType,a.AllocationQty,
                                               a.ManHours,a.Overhead,a.YieldRate,c.ResourceRate,c.OverheadRate,
	                                           b.DepartmentId,b.ProcessId
                                          FROM PDM.DfmQiProcess a
                                               INNER JOIN MES.ProcessParameter b ON a.ParameterId = b.ParameterId
	                                           LEFT JOIN BAS.DepartmentRate c ON b.DepartmentId = c.DepartmentId
                                         WHERE a.DfmQiId = @DfmQiId
                                           AND c.Status = 'A'
                                         ORDER BY a.SortNumber";
                                    dynamicParameters.Add("DfmQiId", DfmQiId);
                                    var DfmQiProcessResult = sqlConnection.Query(sql, dynamicParameters);
                                    foreach (var item2 in DfmQiProcessResult)
                                    {
                                        #region //取出QiProcess參數
                                        int DfmQiProcessId = Convert.ToInt32(item2.DfmQiProcessId);
                                        int SortNumber = Convert.ToInt32(item2.SortNumber);
                                        int AllocationType = Convert.ToInt32(item2.AllocationType);
                                        int AllocationQty = Convert.ToInt32(item2.AllocationQty);
                                        double ManHours = Convert.ToDouble(item2.ManHours);
                                        double MachineHours = Convert.ToDouble(item2.MachineHours);
                                        double ResourceRate = Convert.ToDouble(item2.ResourceRate);
                                        double OverheadRate = Convert.ToDouble(item2.OverheadRate);
                                        #endregion

                                        #region //計算人工費用與製造費用
                                        if (AllocationType.Equals(1))
                                        {
                                            ResourceAmt = ManHours * ResourceRate;
                                            OverheadAmt = MachineHours * OverheadRate;
                                        }
                                        else if (AllocationType.Equals(2))
                                        {
                                            ResourceAmt = ManHours * ResourceRate / AllocationQty;
                                            OverheadAmt = MachineHours * OverheadRate / AllocationQty;
                                        }
                                        else if (AllocationType.Equals(3))
                                        {
                                            ResourceAmt = ManHours * ResourceRate / SolutionQty;
                                            OverheadAmt = MachineHours * OverheadRate / SolutionQty;
                                        }
                                        #endregion

                                        #region //新增DfmQiSolutionProcess
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO PDM.DfmQiSolutionProcess (DfmQiProcessId, DfmQiSolutionId, ResourceAmt, OverheadAmt
                                                ,CreateDate, CreateBy
                                                ,LastModifiedDate, LastModifiedBy)
                                                VALUES (@DfmQiSolutionId, @ResourceAmt, @OverheadAmt
                                                ,@CreateDate, @CreateBy
                                                ,@LastModifiedDate, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                DfmQiProcessId,
                                                DfmQiSolutionId,
                                                ResourceAmt,
                                                OverheadAmt,
                                                CreateDate,
                                                CreateBy = CreateBy,
                                                LastModifiedDate,
                                                LastModifiedBy = CreateBy
                                            });
                                        var insertQisProcess = sqlConnection.Query(sql, dynamicParameters);
                                        #endregion

                                        #region //更新DfmQiSolution的工和費資訊
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE PDM.DfmQiSolution
                                                   SET ResourceAmt = ResourceAmt + @ResourceAmt,
                                                       OverheadAmt = OverheadAmt + @OverheadAmt,
                                                       LastModifiedDate = @LastModifiedDate,
                                                       LastModifiedBy = @LastModifiedBy
                                                 WHERE DfmQiSolutionId = @DfmQiSolutionId";
                                        dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DfmQiSolutionId,
                                            ResourceAmt,
                                            OverheadAmt,
                                            LastModifiedDate = DateTime.Now,
                                            LastModifiedBy = CreateBy
                                        });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                        #endregion
                                    }
                                    #endregion
                                }
                                #endregion
                            }


                            #endregion


                        }

                        if (QuotationResult <= 0) throw new SystemException("沒有可計算之報價項目!");
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion
    }
}
