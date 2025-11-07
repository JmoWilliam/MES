using Dapper;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Transactions;
using System.Web;

namespace SCMDA
{
    public class ScmApiDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";
        public string HrmEtergeConnectionStrings = "";

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

        public ScmApiDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];

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

        #region //PO 採購
        #region//ConfirmQuotation -- 核准核價單(非正式核單邏輯) -- Luca 2024-04-03
        public string ConfirmQuotation(string QoErpPrefix, string QoErpNo, string CompanyNo)
        {
            try
            {
                string ErpDbName = "";
                string ErpNo = "";
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() != 1) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpDbName = ErpConnectionStrings.Split('=')[2].Split(';')[0];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認核價單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TL004)) SupplierNo
                                    , LTRIM(RTRIM(TL005)) Currency
                                    , LTRIM(RTRIM(TL006)) TL006
                                    FROM PURTL
                                    WHERE TL001 = @QoErpPrefix
                                    AND TL002 = @QoErpNo";
                            dynamicParameters.Add("QoErpPrefix", QoErpPrefix);
                            dynamicParameters.Add("QoErpNo", QoErpNo);

                            var PURTCResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (PURTCResult.Count() <= 0) throw new SystemException("查無此核價單據資料!!");

                            string SupplierNo = "";
                            string Currency = "";
                            foreach (var item in PURTCResult)
                            {
                                if (item.TL006 != "N") throw new SystemException("單據狀態不可核單!");
                                SupplierNo = item.SupplierNo;
                                Currency = item.Currency;
                            }
                            #endregion

                            #region //取得單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TM003)) QoSeq, LTRIM(RTRIM(a.TM004)) MtlItemNo, LTRIM(RTRIM(a.TM014)) EffectiveDate
                                    FROM PURTM a 
                                    WHERE a.TM001 = @QoErpPrefix
                                    AND a.TM002 = @QoErpNo";
                            dynamicParameters.Add("QoErpPrefix", QoErpPrefix);
                            dynamicParameters.Add("QoErpNo", QoErpNo);

                            var PURTMResult = sqlConnection2.Query(sql, dynamicParameters);
                            #endregion

                            #region //確認是否有【同廠商】【同品號】【同幣別】【已核單】之其他核價單單身
                            foreach (var item in PURTMResult)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 LTRIM(RTRIM(a.TM001)) QoErpPrefix, LTRIM(RTRIM(a.TM002)) QoErpNo, LTRIM(RTRIM(a.TM003)) QoSeq
                                        , CASE WHEN LEN(LTRIM(RTRIM(TM015))) > 0 THEN CONVERT(VARCHAR, CAST(LTRIM(RTRIM(TM015)) AS DATE), 23) ELSE NULL END AS TM015String
                                        FROM PURTM a 
                                        INNER JOIN PURTL b ON a.TM001 = b.TL001 AND a.TM002 = b.TL002
                                        WHERE a.TM004 = @MtlItemNo
                                        ANd a.TM011 = 'Y'
                                        AND b.TL004 = @SupplierNo
                                        AND b.TL005 = @Currency
                                        AND b.TL006 = 'Y'
                                        AND a.TM015 != ''
                                        AND (a.TM001 + '-' + a.TM002 + '-' + a.TM003) != @QoErpFullNo
                                        UNION ALL 
                                        SELECT LTRIM(RTRIM(a.TM001)) QoErpPrefix, LTRIM(RTRIM(a.TM002)) QoErpNo, LTRIM(RTRIM(a.TM003)) QoSeq
                                        , CASE WHEN LEN(LTRIM(RTRIM(TM015))) > 0 THEN CONVERT(VARCHAR, CAST(LTRIM(RTRIM(TM015)) AS DATE), 23) ELSE NULL END AS TM015String
                                        FROM PURTM a 
                                        INNER JOIN PURTL b ON a.TM001 = b.TL001 AND a.TM002 = b.TL002
                                        WHERE a.TM004 = @MtlItemNo
                                        ANd a.TM011 = 'Y'
                                        AND b.TL004 = @SupplierNo
                                        AND b.TL005 = @Currency
                                        AND b.TL006 = 'Y'
                                        AND a.TM015 = ''
                                        AND (a.TM001 + '-' + a.TM002 + '-' + a.TM003) != @QoErpFullNo
                                        ORDER BY TM015String DESC";
                                dynamicParameters.Add("MtlItemNo", item.MtlItemNo);
                                dynamicParameters.Add("SupplierNo", SupplierNo);
                                dynamicParameters.Add("Currency", Currency);
                                dynamicParameters.Add("QoErpFullNo", QoErpPrefix + "-" + QoErpNo + "-" + item.QoSeq);

                                var CheckPURTMResult = sqlConnection2.Query(sql, dynamicParameters);

                                foreach (var item2 in CheckPURTMResult)
                                {
                                    #region //更新清單內的失效日
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE PURTM SET
                                            TM015 = @EffectiveDate
                                            WHERE TM001 = @QoErpPrefix
                                            AND TM002 = @QoErpNo
                                            AND TM003 = @QoSeq";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          item.EffectiveDate,
                                          item2.QoErpPrefix,
                                          item2.QoErpNo,
                                          item2.QoSeq
                                      });

                                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                            #endregion

                            #region //UPDATE PURTC, PURTD
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTL SET
                                    TL006 = 'Y'
                                    WHERE TL001 = @QoErpPrefix
                                    AND TL002 = @QoErpNo";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  QoErpPrefix,
                                  QoErpNo
                              });

                            rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTM SET
                                    TM011 = 'Y'
                                    WHERE TM001 = @QoErpPrefix
                                    AND TM002 = @QoErpNo";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  QoErpPrefix,
                                  QoErpNo
                              });

                            rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)"
                            });
                            #endregion
                        }
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

        #region //ConfirmPurchaseOrder -- 核准採購單 -- Ann 2023-12-05
        public string ConfirmPurchaseOrder(string PoErpPrefix, string PoErpNo, string CompanyNo)
        {
            try
            {
                string ErpDbName = "";
                string ErpNo = "";
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpDbName = ErpConnectionStrings.Split('=')[2].Split(';')[0];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認採購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TC014)) TC014
                                    FROM PURTC
                                    WHERE TC001 = @PoErpPrefix
                                    AND TC002 = @PoErpNo";
                            dynamicParameters.Add("PoErpPrefix", PoErpPrefix);
                            dynamicParameters.Add("PoErpNo", PoErpNo);

                            var PURTCResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (PURTCResult.Count() <= 0) throw new SystemException("查無此採購單據資料!!");

                            foreach (var item in PURTCResult)
                            {
                                if (item.TC014 != "N") throw new SystemException("單據狀態不可核單!");
                            }
                            #endregion

                            #region //UPDATE PURTC, PURTD
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTC SET
                                    TC014 = 'Y'
                                    WHERE TC001 = @PoErpPrefix
                                    AND TC002 = @PoErpNo";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  PoErpPrefix,
                                  PoErpNo
                              });

                            rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTD SET
                                    TD018 = 'Y'
                                    WHERE TD001 = @PoErpPrefix
                                    AND TD002 = @PoErpNo";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  PoErpPrefix,
                                  PoErpNo
                              });

                            rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)"
                            });
                            #endregion
                        }
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

        #region //ConfirmPurchaseOrderChange -- 核准採購變更單 -- Ann 2024-07-29
        public string ConfirmPurchaseOrderChange(string PoErpPrefix, string PoErpNo, string Edition, string CompanyNo)
        {
            try
            {
                string ErpDbName = "";
                string ErpNo = "";
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpDbName = ErpConnectionStrings.Split('=')[2].Split(';')[0];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認採購變更單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TE017)) TE017
                                    FROM PURTE 
                                    WHERE TE001 = @TE001 
                                    AND TE002 = @TE002 
                                    AND TE003 = @TE003";
                            dynamicParameters.Add("TE001", PoErpPrefix);
                            dynamicParameters.Add("TE002", PoErpNo);
                            dynamicParameters.Add("TE003", Edition);

                            var PURTEResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (PURTEResult.Count() <= 0) throw new SystemException("查無此採購變更單據資料!!");

                            foreach (var item in PURTEResult)
                            {
                                if (item.TE017 != "N") throw new SystemException("單據狀態不可核單!");
                            }
                            #endregion

                            #region //確認是否有比此版本更舊且未確認的變更單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TE017)) TE017
                                    FROM PURTE 
                                    WHERE TE001 = @TE001 
                                    AND TE002 = @TE002 
                                    AND TE003 < @TE003
                                    AND TE017 != 'Y'";
                            dynamicParameters.Add("TE001", PoErpPrefix);
                            dynamicParameters.Add("TE002", PoErpNo);
                            dynamicParameters.Add("TE003", Edition);

                            var PURTECheckResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (PURTECheckResult.Count() > 0) throw new SystemException("此採購單存在版本更小的變更單且未核單!!");
                            #endregion

                            #region //UPDATE PURTE, PURTF
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTE SET
                                    TC017 = 'Y'
                                    WHERE TE001 = @PoErpPrefix
                                    AND TE002 = @PoErpNo
                                    AND TE003 = @Edition";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  PoErpPrefix,
                                  PoErpNo,
                                  Edition
                              });

                            rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTF SET
                                    TF016 = 'Y'
                                    WHERE TF001 = @PoErpPrefix
                                    AND TF002 = @PoErpNo
                                    AND TF003 = @Edition";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  PoErpPrefix,
                                  PoErpNo,
                                  Edition
                              });

                            rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)"
                            });
                            #endregion
                        }
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
        #endregion
    }
}
