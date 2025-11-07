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

namespace SCMDA
{
    public class TakeOrderDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string ErpSysConnectionStrings = "";
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

        public TakeOrderDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpSysConnectionStrings = ConfigurationManager.AppSettings["ErpSysDb"];
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
        #region//AddDemand  --需求單單頭 新增 -- Ding 2022.12.26
        public string AddDemand(string DemandNo, string CompanyNo, string DemandSource, string DemandDesc, string CustNo,
            string Remark, string DemandDate)
        {
            try
            {
                if (DemandNo == "") throw new SystemException("【需求單編號】不能為空!");
                if (CompanyNo == "") throw new SystemException("【公司編號】不能為空!");
                if (DemandSource == "") throw new SystemException("【單據來源】不能為空!");
                if (CustNo == "") throw new SystemException("【客戶編號】不能為空!");
                if (DemandDate == "") throw new SystemException("【提出日】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //Get CompanyId
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId
                            FROM BAS.Company a                              
                            WHERE a.MoId =@MoId
                            AND a.CompanyNo=@CompanyNo                                
                            ";
                        dynamicParameters.Add("CompanyNo", CompanyNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        int CompanyId = -1;
                        foreach (var item in result)
                        {
                            CompanyId = item.CompanyId;
                        }
                        #endregion

                        #region //Get Customer
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId,a.CustomerNo
                            FROM SCM.Customer a                              
                            WHERE a.CustomerNo =@CustomerNo
                            AND a.CompanyId=@CompanyId                                
                            ";
                        dynamicParameters.Add("CompanyId", CompanyId);
                        dynamicParameters.Add("CustomerNo", CustNo);
                        var result1 = sqlConnection.Query(sql, dynamicParameters);
                        if (result1.Count() > 0) throw new SystemException("客戶編號】查無資料!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.Demand (CompanyId,DemandNo,DemandSource,DemandDesc,CustNo,Remark,DemandDate
                            ,CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                            OUTPUT INSERTED.DemandId
                            VALUES (@CompanyId, @DemandNo, @DemandSource,@DemandDesc,@CustNo,@Remark,@DemandDate
                            ,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId,
                                DemandNo,
                                DemandSource,
                                DemandDesc,
                                CustNo,
                                Remark,
                                DemandDate,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        transactionScope.Complete();
                    }
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

        #region//AddDemandLine  --需求單單身 新增 -- Ding 2022.12.26
        public string AddDemandLine(string DemandLineNo, int DemandId, int ProductTypeId, string TypeNo,
            string WipCategory, string CustMtlName, string DeliveryDate, int UnitPrice,
            string UomNo, int OrderQty, string OrderType, string MtlSpec, string DeliveryProcess, string Remark)
        {
            try
            {
                if (DemandLineNo == "") throw new SystemException("【需求單編號】不能為空!");
                if (DemandId < 0) throw new SystemException("【需求單編號】不能為空!");
                if (ProductTypeId < 0) throw new SystemException("【產品結構類別】不能為空!");
                if (TypeNo == "") throw new SystemException("【產品屬性】不能為空!");
                if (WipCategory == "") throw new SystemException("【製令單別】不能為空!");
                if (CustMtlName == "") throw new SystemException("【客戶部番】不能為空!");
                if (DeliveryDate == "") throw new SystemException("【要望交期】不能為空!");
                if (OrderType == "") throw new SystemException("【首/複製番】不能為空!");

                #region//定義值
                int CompanyId = -1;
                #endregion


                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region//Demand 需求單所屬公司
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId
                            FROM MES.Demand a
                            WHERE a.DemandId = @DemandId";
                        dynamicParameters.Add("DemandId", DemandId);
                        var Result = sqlConnection.Query(sql, dynamicParameters);

                        if (Result.Count() <= 0) throw new SystemException("查無需求單!");

                        foreach (var item in Result)
                        {
                            CompanyId = int.Parse(item.CompanyId);
                        }
                        #endregion

                        #region //Get ProdTypeId
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ProductTypeId
                            FROM MES.DemandProductType a                              
                            WHERE a.ProductTypeId =@ProductTypeId
                            AND a.CompanyId=@CompanyId";
                        dynamicParameters.Add("ProductTypeId", ProductTypeId);
                        dynamicParameters.Add("CompanyId", CompanyId);
                        var result1 = sqlConnection.Query(sql, dynamicParameters);
                        if (result1.Count() != 1) throw new SystemException("【產品結構類別】查無資料!");
                        #endregion

                        #region //Get TypeNo
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TypeNo
                            FROM BAS.[Type] a                              
                            WHERE a.TypeSchema ='DemandTypeNo'
                            AND a.TypeNo=@TypeNo";
                        dynamicParameters.Add("TypeNo", TypeNo);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() != 1) throw new SystemException("【產品屬性】查無資料!");
                        #endregion

                        #region //Get DemandOrderType
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TypeNo
                            FROM BAS.[Type] a                              
                            WHERE a.TypeSchema ='DemandOrderType'
                            AND a.TypeNo=@TypeNo";
                        dynamicParameters.Add("OrderType", OrderType);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() != 1) throw new SystemException("【首/複製番】查無資料!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.Demand (DemandLineNo,DemandId,ProductTypeId,TypeNo,WipCategory,CustMtlName
                            ,DeliveryDate,UnitPrice,UomNo,OrderQty,OrderType,MtlSpec,DeliveryProcess,Remark
                            ,CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                            OUTPUT INSERTED.DemandId
                            VALUES (@CompanyId, @DemandNo, @DemandSource,@DemandDesc,@CustNo,@Remark,@DemandDate
                            ,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DemandLineNo,
                                DemandId,
                                ProductTypeId,
                                TypeNo,
                                WipCategory,
                                CustMtlName,
                                DeliveryDate,
                                UnitPrice,
                                UomNo,
                                OrderQty,
                                OrderType,
                                MtlSpec,
                                DeliveryProcess,
                                Remark,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        transactionScope.Complete();
                    }
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
        #endregion

        #region//Update
        #endregion

        #region//Delete
        #endregion

    }
}
