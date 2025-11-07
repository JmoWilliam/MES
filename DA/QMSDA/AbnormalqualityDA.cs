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
using System.Transactions;
using System.Web;

namespace QMSDA
{
    public class AbnormalqualityDA
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

        public AbnormalqualityDA()
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

        #region //Get

        #region //GetAbnormalquality -- 取得品異單單頭 -- Shintokuro 2022.10.06
        public string GetAbnormalquality(int AbnormalqualityId, string AbnormalqualityNo, string CreateDate, string AbnormalqualityStatus
            , string StartCreateDate, string EndCreateDate, string BarcodeNo,string ErpNo,int MoId, int ResponsibleDeptId
            , string AqBarcodeStatus, string JudgeConfirm
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                List<dynamic> dataList2 =new List<dynamic>();

                if(ResponsibleDeptId > 0)
                {
                    PageSize = 9999;
                }

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.AbnormalqualityId";
                    sqlQuery.auxKey = ",b.DepartmentId";
                    sqlQuery.columns =
                        @",a.AbnormalqualityNo, a.AbnormalqualityStatus,FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate, a.CompanyId,a.QcType
                          ,b.ProcessAlias, b.MoProcessId
                          , c.MoId,(c1.WoErpPrefix + '-' + c1.WoErpNo + '(' + CONVERT ( Varchar , c.WoSeq ) + ')') MoNo
                          ,c2.MtlItemNo ,c2.MtlItemName
                          ,d.GrDetailId,d.LotNumber
                          ,(d1.GrErpPrefix + '-' + d1.GrErpNo + '(' + CONVERT ( Varchar(4) , d.GrSeq ) + ')') GrErpFull
                          ,d2.MtlItemNo GrMtlItemNo ,d2.MtlItemName GrMtlItemName
                          , (
                             SELECT x.ResponsibleDeptId,x.ResponsibleUserId,x.AqBarcodeStatus,x.JudgeConfirm
                             ,x.AqBarcodeId,x1.BarcodeNo
                             ,y.DepartmentId, y.DepartmentName,y1.UserNo,y1.UserName
                             FROM QMS.AqBarcode x
                             LEFT JOIN MES.Barcode x1 on x.BarcodeId = x1.BarcodeId
                             LEFT JOIN BAS.Department y ON x.ResponsibleDeptId = y.DepartmentId
                             LEFT JOIN BAS.[User] y1 ON x.ResponsibleUserId = y1.UserId
                             WHERE x.AbnormalqualityId = a.AbnormalqualityId
                             FOR JSON PATH, ROOT('data')
                          ) AqBarcode
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.Abnormalquality a
                          LEFT JOIN MES.MoProcess b on a.MoProcessId = b.MoProcessId
                          LEFT JOIN MES.ManufactureOrder c on a.MoId = c.MoId
                          LEFT JOIN MES.WipOrder c1 on c.WoId = c1.WoId
                          LEFT JOIN PDM.MtlItem c2 on c1.MtlItemId = c2.MtlItemId
                          LEFT JOIN SCM.GrDetail d on a.GrDetailId = d.GrDetailId
                          LEFT JOIN SCM.GoodsReceipt d1 on d.GrId = d1.GrId
                          LEFT JOIN PDM.MtlItem d2 on d.MtlItemId = d2.MtlItemId
                          --LEFT JOIN QMS.AqBarcode d ON d.AbnormalqualityId = a.AbnormalqualityId
                          --LEFT JOIN BAS.Department d1 ON d1.DepartmentId = d.ResponsibleDeptId
                         ";
                    string queryTable = @"FROM (
                            SELECT a.AbnormalqualityId,a.AbnormalqualityNo, a.AbnormalqualityStatus,FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                                   , a.CompanyId,a.QcType
                                   ,b.ProcessAlias, b.MoProcessId
                                   , c.MoId,(c1.WoErpPrefix + '-' + c1.WoErpNo + '(' + CONVERT ( Varchar , c.WoSeq ) + ')') MoNo
                                   ,c2.MtlItemNo ,c2.MtlItemName
                                   ,(
                                       SELECT x.ResponsibleDeptId,x.ResponsibleUserId,x.AqBarcodeStatus,x.JudgeConfirm
                                       ,x.AqBarcodeId,x1.BarcodeNo
                                       ,y.DepartmentId, y.DepartmentName,y1.UserNo,y1.UserName
                                       FROM QMS.AqBarcode x
                                       LEFT JOIN MES.Barcode x1 on x.BarcodeId = x1.BarcodeId
                                       LEFT JOIN BAS.Department y ON x.ResponsibleDeptId = y.DepartmentId
                                       LEFT JOIN BAS.[User] y1 ON x.ResponsibleUserId = y1.UserId
                                       WHERE x.AbnormalqualityId = a.AbnormalqualityId
                                       FOR JSON PATH, ROOT('data')
                                   ) AqBarcode
                            FROM QMS.Abnormalquality a
                            LEFT JOIN MES.MoProcess b on a.MoProcessId = b.MoProcessId
                            LEFT JOIN MES.ManufactureOrder c on a.MoId = c.MoId
                            LEFT JOIN MES.WipOrder c1 on c.WoId = c1.WoId
                            LEFT JOIN PDM.MtlItem c2 on c1.MtlItemId = c2.MtlItemId
                        ) a  
                        OUTER APPLY(
                            SELECT(
                               SELECT CAST( x.DepartmentId AS nvarChar) + ','
                               FROM OPENJSON(a.AqBarcode, '$.data')
                               WITH(
                                    DepartmentId INT N'$.DepartmentId'
                               ) x
                               FOR XML PATH('')
                            ) DepartmentId
                        ) b";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AbnormalqualityId", @" AND a.AbnormalqualityId = @AbnormalqualityId", AbnormalqualityId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AbnormalqualityNo", @" AND a.AbnormalqualityNo LIKE '%' + @AbnormalqualityNo + '%'", AbnormalqualityNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AbnormalqualityStatus", @" AND a.AbnormalqualityStatus = @AbnormalqualityStatus", AbnormalqualityStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartCreateDate", @" AND a.CreateDate >= @StartCreateDate ", StartCreateDate.Length > 0 ? Convert.ToDateTime(StartCreateDate).ToString("yyyy-MM-dd 00:00:000") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndCreateDate", @" AND a.CreateDate <= @EndCreateDate ", EndCreateDate.Length > 0 ? Convert.ToDateTime(EndCreateDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BarcodeNo", @" AND exists(select 1 from QMS.AqBarcode b inner join MES.Barcode c ON b.BarcodeId = c.BarcodeId where a.AbnormalqualityId = b.AbnormalqualityId and c.BarcodeNo = @BarcodeNo)", BarcodeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ErpNo", @" AND a.MoNo LIKE '%'+ @ErpNo+ '%' ", ErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MoId", @" AND a.MoId = @MoId", MoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AqBarcodeStatus", @" AND EXISTS ( SELECT x.AqBarcodeStatus FROM QMS.AqBarcode x WHERE a.AbnormalqualityId = x.AbnormalqualityId AND x.AqBarcodeStatus = @AqBarcodeStatus)", AqBarcodeStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "JudgeConfirm", @" AND EXISTS ( SELECT x.JudgeConfirm FROM QMS.AqBarcode x WHERE a.AbnormalqualityId = x.AbnormalqualityId AND x.JudgeConfirm = @JudgeConfirm)", JudgeConfirm);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ErpNo", @" AND (a.WoErpPrefix + '-' + a.WoErpNo + '(' + CONVERT ( Varchar , a.WoSeq ) + ')') LIKE '%'+ @ErpNo+ '%' ", ErpNo);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ResponsibleDeptId", @" AND b.DepartmentId = @ResponsibleDeptId", ResponsibleDeptId);
                    //string testRR = Convert.ToDateTime(StartCreateDate).ToString("yyyy-MM-dd 00:00:000");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.AbnormalqualityId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);
                    if (ResponsibleDeptId > 0) //部門條件篩選
                    {
                        List<int> DepartmentList = new List<int>();
                        List<dynamic> dataList = result.ToList();
                        int place = 0;
                        foreach (var item in result)
                        {
                            if(item.DepartmentId != null)
                            {
                                string mate = "N";

                                var DepartmentIdStr = item.DepartmentId;
                                var DepartmentIdStr2 = DepartmentIdStr.Split(',');
                                foreach (var item2 in DepartmentIdStr2)
                                {
                                    if (item2 == ResponsibleDeptId.ToString())
                                    {
                                        if (!DepartmentList.Contains(item.AbnormalqualityId))
                                        {
                                            DepartmentList.Add(item.AbnormalqualityId);
                                            mate = "Y";
                                        }
                                    }
                                }
                                if(mate != "Y")
                                {
                                    dataList[place] = null;
                                }
                            }
                            else
                            {
                                dataList[place] = null;
                            }
                            place++;
                        }

                        dataList.RemoveAll(item => item == null);
                        dataList2 = dataList;
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = ResponsibleDeptId>0 ? dataList2 : result
                    });
                    #endregion
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

        #region //GetAqDetail -- 取得品異條碼單身 -- Shintokuro 2022.10.07
        public string GetAqDetail(int AqBarcodeId, int AbnormalqualityId, string AqBarcodeStatus, string JudgeConfirm
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.AqBarcodeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" ,a.AbnormalqualityId, a.QcBarcodeId, a.DefectCauseId, a.DefectCauseDesc, a.ConformUserId
                           ,a.ResponsibleDeptId, a.ResponsibleUserId, a.SubResponsibleDeptId, a.SubResponsibleUserId, a.ProgrammerUserId
                           ,a.RepairCauseId, a.RepairCauseDesc, a.RepairCauseUserId, a.ResponsibleSupervisorId, a.JudgeStatus, a.JudgeDesc, a.JudgeUserId
                           ,a.JudgeReturnMoProcessId ,a.JudgeReturnNextMoProcessId ,a.CountersignUserId, a.AqBarcodeStatus, a.CreateDate, a.LastModifiedDate, a.CreateBy
                           ,a.LastModifiedBy, a.JudgeConfirm,FORMAT(a.JudgeDate, 'yyyy-MM-dd') JudgeDate, a.ChangeJudgeFlag,a.SupplierId,a.ReleaseQty
                           ,a1.MoId,a1.AbnormalqualityNo ,a1.AbnormalqualityStatus ,FORMAT(a1.CreateDate, 'yyyy-MM-dd') AqrCreateDate
                           ,a1.QcType, a2.TypeName
                           ,b.ProcessAlias
                           ,b1.InputQty
                           ,(b2.WoErpPrefix + '-' + b2.WoErpNo + '(' + CONVERT ( Varchar , b1.WoSeq ) + ')') MoNo ,b3.MtlItemNo ,b3.MtlItemName
                           ,b4.BarcodeNo,b4.BarcodeQty
                           ,b6.ModeId ,b6.ModeName
                           ,b7.QcStatus ,b7.QcUserId,b7.Remark
                           ,c.CauseNo  DefectCauseNo ,c.CauseName DefectCauseName
                           ,d.CauseNo  RepairCauseNo ,d.CauseName RepairCauseName
                           ,e1.UserName ConformUserName ,e1.Gender ConformUserGender
                           ,e2.UserName ResponsibleUserName ,e2.UserNo ResponsibleUserNo ,e2.Gender ResponsibleUserGender
                           ,e3.UserName SubResponsibleUserName ,e3.Gender SubResponsibleUserGender
                           ,e4.UserName ProgrammerUserName ,e4.Gender ProgrammerUserGender
                           ,e5.UserName RepairCauseUserName ,e5.UserNo RepairCauseUserNo ,e5.Gender RepairCauseUserGender
                           ,e6.UserName ResponsibleSupervisorName ,e6.UserNo ResponsibleSupervisorNo ,e6.Gender ResponsibleSupervisorGender
                           ,e7.UserName JudgeUserName ,e7.UserNo JudgeUserNo ,e7.Gender JudgeUserGender
                           ,e8.UserName CountersignUserName ,e8.Gender CountersignUserGender
                           ,f.DepartmentName ResponsibleDepartmentName,g.DepartmentName SubResponsibleDepartmentName
                           ,h.ProcessAlias JudgeReturnMoProcessAlias
                           ,h1.ProcessAlias JudgeReturnNextMoProcessAlias
                           ,i.GrDetailId,i.LotNumber GrLotNumber
                           ,(i1.GrErpPrefix + '-' + i1.GrErpNo + '(' + CONVERT ( Varchar(4) , i.GrSeq ) + ')') GrErpFull
                           ,i2.MtlItemNo GrMtlItemNo ,i2.MtlItemName GrMtlItemName
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.AqBarcode a
                          INNER JOIN QMS.Abnormalquality a1 on a.AbnormalqualityId = a1.AbnormalqualityId
                          INNER JOIN BAS.[Type] a2 on a1.QcType = a2.TypeNo AND TypeSchema ='QcItem.QcType'
                          LEFT JOIN MES.MoProcess b on a1.MoProcessId = b.MoProcessId
                          LEFT JOIN MES.ManufactureOrder b1 on b.MoId = b1.MoId
                          LEFT JOIN MES.WipOrder b2 on b1.WoId = b2.WoId
                          LEFT JOIN PDM.MtlItem b3 on b2.MtlItemId = b3.MtlItemId
                          LEFT JOIN MES.Barcode b4 on a.BarcodeId = b4.BarcodeId
                          LEFT JOIN MES.ProdMode b6 on b1.ModeId = b6.ModeId
                          LEFT JOIN MES.QcBarcode b7 on a.QcBarcodeId = b7.QcBarcodeId
                          LEFT JOIN QMS.DefectCause c on a.DefectCauseId = c.CauseId
                          LEFT JOIN QMS.RepairCause d on a.RepairCauseId = d.CauseId
                          LEFT JOIN BAS.[User] e1 on a.ConformUserId = e1.UserId
                          LEFT JOIN BAS.[User] e2 on a.ResponsibleUserId = e2.UserId
                          LEFT JOIN BAS.[User] e3 on a.SubResponsibleUserId = e3.UserId
                          LEFT JOIN BAS.[User] e4 on a.ProgrammerUserId = e4.UserId
                          LEFT JOIN BAS.[User] e5 on a.RepairCauseUserId = e5.UserId
                          LEFT JOIN BAS.[User] e6 on a.ResponsibleSupervisorId = e6.UserId
                          LEFT JOIN BAS.[User] e7 on a.JudgeUserId = e7.UserId
                          LEFT JOIN BAS.[User] e8 on a.CountersignUserId = e8.UserId
                          LEFT JOIN BAS.Department f on a.ResponsibleDeptId = f.DepartmentId
                          LEFT JOIN BAS.Department g on a.SubResponsibleDeptId = g.DepartmentId
                          LEFT JOIN MES.MoProcess h on a.JudgeReturnMoProcessId = h.MoProcessId
                          LEFT JOIN MES.MoProcess h1 on a.JudgeReturnNextMoProcessId = h1.MoProcessId
                          LEFT JOIN SCM.GrDetail i on a1.GrDetailId = i.GrDetailId
                          LEFT JOIN SCM.GoodsReceipt i1 on i.GrId = i1.GrId
                          LEFT JOIN PDM.MtlItem i2 on i.MtlItemId = i2.MtlItemId
                          ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AqBarcodeId", @" AND a.AqBarcodeId = @AqBarcodeId", AqBarcodeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AbnormalqualityId", @" AND a.AbnormalqualityId = @AbnormalqualityId", AbnormalqualityId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "JudgeConfirm", @" AND a.JudgeConfirm = @JudgeConfirm", JudgeConfirm);
                    if(AqBarcodeStatus !="") BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AqBarcodeStatus", @" AND a.AqBarcodeStatus in @AqBarcodeStatus", AqBarcodeStatus.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.AqBarcodeId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
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

        #region //GetAqQcItem -- 取得品異條碼異常項目原因 -- Shintokuro 2022.10.14
        public string GetAqQcItem(int AqBarcodeId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.AqQcItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.AqBarcodeId, a.QcItemId, a.DefectCauseId, a.DefectCauseId, a.DefectCauseDesc, a.RepairCauseId, a.RepairCauseDesc
                          ,b.CauseNo DefectCauseNo,b.CauseName DefectCauseName
                          ,c.QcItemName
                          ,d.CauseNo RepairCauseNo,d.CauseName RepairCauseName
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.AqQcItem a
                          INNER JOIN QMS.DefectCause b on a.DefectCauseId = b.CauseId
                          INNER JOIN QMS.QcItem c on a.QcItemId = c.QcItemId
                          LEFT JOIN QMS.RepairCause d on a.RepairCauseId = d.CauseId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @" AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AqBarcodeId", @" AND a.AqBarcodeId = @AqBarcodeId", AqBarcodeId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.AqQcItemId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
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

        #region //GetAqFile -- 取得品異條碼原因佐證檔案 -- Shintokuro 2022.10.14
        public string GetAqFile(int AqFileId, int AqBarcodeId, string AqFileStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.AqFileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.AqBarcodeId, a.AqFileStatus, a.FileId
                           ,b.[FileName]
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.AqFile a
                          INNER JOIN BAS.[File] b on a.FileId = b.FileId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AqBarcodeId", @" AND a.AqBarcodeId = @AqBarcodeId", AqBarcodeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AqFileStatus", @" AND a.AqFileStatus = @AqFileStatus", AqFileStatus);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.AqFileId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
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

        #region //GetAqFileShow -- 取得品異條碼原因佐證檔案(顯示) -- Shintokuro 2022.10.17
        public string GetAqFileShow(int AqBarcodeId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.AqFileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.AqBarcodeId, a.AqFileStatus, a.FileId
                           ,b.[FileName]
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.AqFile a
                          INNER JOIN BAS.[File] b on a.FileId = b.FileId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AqBarcodeId", @" AND a.AqBarcodeId = @AqBarcodeId", AqBarcodeId);


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.AqFileId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
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

        #region //GetJudgeReturnMoProcess -- 取得判定退回站別清單 -- Shintokuro 2022.10.17
        public string GetJudgeReturnMoProcess(int MoId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MoProcessId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SortNumber, a.ProcessAlias, a.MoId
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.MoProcess a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MoId", @" AND a.MoId = @MoId", MoId);


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SortNumber ASC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
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

        #region //GetBarcodeList -- 取得條碼列表 -- Shintokuro 2022.10.20
        public string GetBarcodeList(int MoId, int MoProcessId, string BarcodeListView
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.auxKey = "";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = "";
                    switch (BarcodeListView)
                    {

                        case "OQC":
                            sqlQuery.mainKey = "a.BarcodeId";

                            sqlQuery.columns =
                                @", a.BarcodeNo,a.BarcodeStatus
                                ";
                            sqlQuery.mainTables =
                                @" FROM MES.Barcode a
                                 ";
                            queryCondition = @"AND a.BarcodeId not in (select  BarcodeId FROM QMS.AqBarcode b WHERE JudgeConfirm !='Y' AND b.BarcodeId = a.BarcodeId) AND a.NextMoProcessId = -1 AND a.BarcodeStatus = 0";
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MoId", @" AND a.MoId = @MoId", MoId);
                            sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.BarcodeId ASC";

                            sqlQuery.conditions = queryCondition;
                            sqlQuery.pageIndex = PageIndex;
                            sqlQuery.pageSize = PageSize;

                            var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = result
                            });
                            #endregion
                            break;
                        case "IPQC":
                            sqlQuery.mainKey = "a.BarcodeId";

                            sqlQuery.columns =
                                @", a.BarcodeNo,a.BarcodeStatus
                                  ,c.StartDate,c.FinishDate 
                                ";
                            sqlQuery.mainTables =
                                @" FROM MES.Barcode a
		                           LEFT JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
                                   OUTER  APPLY(
	                                    SELECT TOP 1 c1.StartDate,c1.FinishDate  from MES.BarcodeProcess c1
	                                    WHERE b.MoProcessId = c1.MoProcessId
                                        AND a.BarcodeId = c1.BarcodeId
                                   ) c
                                 ";
                            queryCondition = @"AND a.BarcodeId not in (select  BarcodeId FROM QMS.AqBarcode b WHERE JudgeConfirm !='Y' AND b.BarcodeId = a.BarcodeId) AND　c.FinishDate is not null";
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CurrentMoProcessId", @" AND a.CurrentMoProcessId = @CurrentMoProcessId", MoProcessId);
                            sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.BarcodeId ASC";

                            sqlQuery.conditions = queryCondition;
                            sqlQuery.pageIndex = PageIndex;
                            sqlQuery.pageSize = PageSize;

                            var result1 = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = result1
                            });
                            #endregion
                            break;
                        case "MPC":
                            sql = @"SELECT a.MoId
                                   ,b1.BarcodeRegisterId ,b1.BarcodeNo MrBarcodeNo
                                   ,c1.PrintId ,c1.BarcodeNo PrBarcodeNo
                                   FROM MES.ManufactureOrder a
                                   LEFT JOIN MES.MrDetail b on a.MoId = b.MoId
                                   LEFT JOIN MES.MrBarcodeRegister b1 on b.MrDetailId = b1.MrDetailId
                                   LEFT JOIN MES.MoSetting c on a.MoId =c.MoId
                                   LEFT JOIN MES.BarcodePrint c1 on c.MoSettingId =c1.MoSettingId
                                   WHERE a.MoId = @MoId
                                   ORDER BY b1.BarcodeRegisterId　, c1.PrintId　ASC
                                ";
                            dynamicParameters.Add("MoId", MoId);
                            var result2 = sqlConnection.Query(sql, dynamicParameters);

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = result2
                            });
                            #endregion

                            break;
                        default:
                            break;
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

        #region //GetBarcodeLettering -- 獲取刻字條碼 -- Ted 2022-09-15
        public string GetBarcodeLettering(string BarcodeNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    if (BarcodeNo.Length <= 0) throw new SystemException("【品異單條碼檢視】缺少條碼!!!");

                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.BarcodeNo,b.ItemValue FROM MES.Barcode a
                            LEFT JOIN MES.BarcodeAttribute b on a.BarcodeId = b.BarcodeId
                            WHERE b.ItemNo = 'Lettering'
                            AND a.BarcodeNo = @BarcodeNo";
                    dynamicParameters.Add("BarcodeNo", BarcodeNo);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
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

        #region //GetAqPhrase -- 取得常用片語-- Shintokuro 2024.06.27
        public string GetAqPhrase(string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.AqpId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.PhraseText
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.AqPhrase a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.AqpId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
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

        #region //Add

        #region //AddAbnormalquality-- 新增品異單單頭 -- Shintokuro 2022-10-20
        public string AddAbnormalquality(int? MoId,int? GrDetailId, int? MoProcessId,string QcType, string DocDate, int ViewCompanyId)
        {
            try
            {
                if (ViewCompanyId != CurrentCompany) throw new SystemException("頁面公司別與系統後台公司別不相同,請重新登入系統");
                if (QcType.Length <= 0) throw new SystemException("【檢驗類別】不能為空!");
                if (QcType != "IQC")
                {
                    if (MoId <= 0) throw new SystemException("【製令】不能為空!");
                    GrDetailId = null;
                }
                else
                {
                    if (GrDetailId <= 0) throw new SystemException("【進貨單】不能為空!");
                }

                if (DocDate.Length <= 0) throw new SystemException("【單據日期】不能為空!");


                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //資料驗證
                        if (QcType != "IQC")
                        {
                            #region //判斷製令資料是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.CompanyId
                                FROM MES.ManufactureOrder a
                                INNER JOIN MES.WipOrder b on a.WoId = b.WoId
                                INNER JOIN MES.ProdMode c on a.ModeId = c.ModeId
                                WHERE a.MoId = @MoId";
                            dynamicParameters.Add("MoId", MoId);
                            var resultMoId = sqlConnection.Query(sql, dynamicParameters);
                            if (resultMoId.Count() <= 0) throw new SystemException("製令資料錯誤或不存在!");
                            foreach (var item in resultMoId)
                            {
                                if (item.CompanyId != CurrentCompany) throw new SystemException("製令資料的公司別與後端資料不相符,請重新登入再繼續操作");
                            }
                            #endregion
                        }
                        else
                        {
                            #region //判斷進貨單資料是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.CompanyId,a.ConfirmStatus,a.TransferStatus
                                    FROM SCM.GoodsReceipt a
                                    INNER JOIN SCM.GrDetail b on a.GrId = b.GrId
                                    WHERE b.GrDetailId = @GrDetailId";
                            dynamicParameters.Add("GrDetailId", GrDetailId);
                            var resultGrId = sqlConnection.Query(sql, dynamicParameters);
                            if (resultGrId.Count() <= 0) throw new SystemException("進貨單資料錯誤或不存在!");
                            foreach (var item in resultGrId)
                            {
                                if(item.ConfirmStatus != "N") throw new SystemException("進貨單未確認狀態下才能異動");
                                if(item.TransferStatus != "N") throw new SystemException("進貨單未拋轉狀態下才能異動");
                                if (item.CompanyId != CurrentCompany) throw new SystemException("進貨單資料的公司別與後端資料不相符,請重新登入再繼續操作");
                            }
                            #endregion
                        }
                        #endregion

                        #region //QcType操作
                        switch (QcType)
                        {
                            case "IPQC":
                                if (MoProcessId <= 0) throw new SystemException("檢驗類別為工程檢時,【製程】不能為空!");
                                break;
                            case "OQC":
                                MoProcessId = null;

                                #region //判斷該製令是否未完成判定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT JudgeConfirm
                                    FROM QMS.Abnormalquality a
                                    INNER JOIN QMS.AqBarcode b on a.AbnormalqualityId = b.AbnormalqualityId
								    WHERE a.MoId = @MoId
								    AND b.JudgeConfirm != 'Y'";
                                dynamicParameters.Add("MoId", MoId);
                                var result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 1) throw new SystemException("該製令目前在品異單判定中");
                                #endregion
                                break;
                            case "IQC":
                                MoId = null;
                                MoProcessId = null;
                                #region //判斷進貨單資料是否完成判定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT JudgeConfirm
                                    FROM QMS.Abnormalquality a
                                    INNER JOIN QMS.AqBarcode b on a.AbnormalqualityId = b.AbnormalqualityId
								    WHERE a.GrDetailId = @GrDetailId
								    AND b.JudgeConfirm != 'Y'";
                                dynamicParameters.Add("GrDetailId", GrDetailId);
                                var resultGrDetailId = sqlConnection.Query(sql, dynamicParameters);
                                if (resultGrDetailId.Count() > 0) throw new SystemException("該進貨單目前在品異單判定中");
                                #endregion
                                break;
                            case "MFGQC":
                                break;
                            default:
                                throw new SystemException("該檢驗類別尚未規劃品異");
                                break;
                        }
                        #endregion

                        #region //單號自動取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(AbnormalqualityNo))), '000'), 3)) + 1 CurrentNum
                                FROM QMS.Abnormalquality
								WHERE AbnormalqualityNo NOT LIKE '%[A-Za-z]%'";
                        dynamicParameters.Add("CreateDate", CreateDate);

                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                        string DocDateNo = CreateDate.ToString("yyyyMM");
                        string AbnormalqualityNo = DocDateNo + string.Format("{0:000}", currentNum);
                        #endregion

                        #region //品異單單頭建立
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.Abnormalquality (CompanyId, MoId, GrDetailId, MoProcessId, AbnormalqualityNo, AbnormalqualityStatus, DocDate, QcType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.AbnormalqualityId,INSERTED.AbnormalqualityNo,INSERTED.AbnormalqualityStatus,INSERTED.QcType
                                VALUES (@CompanyId, @MoId, @GrDetailId, @MoProcessId, @AbnormalqualityNo, @AbnormalqualityStatus, @DocDate, @QcType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                MoId,
                                GrDetailId,
                                MoProcessId,
                                AbnormalqualityNo,
                                AbnormalqualityStatus = "F",
                                DocDate,
                                QcType,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        rowsAffected += insertResult.Count();
                        
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

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

        #region //AddAqDetail-- 新增品異單單身 - WEB版本 -- Shintokuro 2022-10-20
        public string AddAqDetail(int AbnormalqualityId , string BarcodeIds, int DefectCauseId, string DefectCauseDesc
            , int ResponsibleDeptId, int ResponsibleUserId, int SubResponsibleDeptId, int SubResponsibleUserId
            , int ProgrammerUserId, int ConformUserId, int? RepairCauseId, string RepairCauseDesc, int? RepairCauseUserId
            , int SupplierId, string ChangeJudgeFlag)
        {
            try
            {

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        int? nullData = null;
                        int MoProcessId = 0;
                        int? GrDetailId = null;
                        string AqBarcodeStatus = "1";
                        //var RepairCauseDesc = nullData;
                        var ResponsibleSupervisorId = nullData;

                        if (DefectCauseId <= 0) throw new SystemException("【不良代碼】不能為空!");
                        if (ResponsibleDeptId <= 0) throw new SystemException("【責任單位】不能為空!");
                        if (ResponsibleUserId <= 0) throw new SystemException("【責任者】不能為空!");
                        if (DefectCauseDesc.ToString().Length > 100) throw new SystemException("【不良原因】長度錯誤!");
                        if (RepairCauseDesc.ToString().Length > 100) throw new SystemException("【對策原因】長度錯誤!");

                        #region //判斷品異單是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MoProcessId,a.QcType,a.GrDetailId
                                FROM QMS.Abnormalquality a
                                WHERE a.AbnormalqualityId = @AbnormalqualityId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【品異單】不存在，請重新輸入!");
                        string QcType = "";
                        foreach (var item in result)
                        {
                            QcType = item.QcType;
                            if (QcType == "IPQC")
                            {
                                MoProcessId = item.MoProcessId;
                            }
                            if (QcType != "IQC")
                            {
                                if (BarcodeIds.Length <= 0) throw new SystemException("【條碼ID】不能為空!");
                            }
                            else
                            {
                                GrDetailId = item.GrDetailId;
                            }
                        }
                        #endregion

                        #region //判斷資料是否存在
                        #region //資料 - NOT NULL
                        #region //判斷不良原因代碼是否存在
                        dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                    FROM QMS.DefectCause a
                                    WHERE a.CauseId = @DefectCauseId";
                            dynamicParameters.Add("DefectCauseId", DefectCauseId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【不良原因代碼】不存在，請重新輸入!");
                            #endregion

                        #region //判斷責任單位是否存在
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b on a.DepartmentId = b.DepartmentId
                                WHERE b.DepartmentId = @ResponsibleDeptId";
                        dynamicParameters.Add("ResponsibleDeptId", ResponsibleDeptId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【責任單位】不存在，請重新輸入!");
                        #endregion

                        #region //判斷責任者是否存在
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ResponsibleUserId
                                AND a.Status = 'A'";
                        dynamicParameters.Add("ResponsibleUserId", ResponsibleUserId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【責任者】不存在，請重新輸入!");
                        #endregion
                        #endregion

                        #region //資料 - NULL
                        #region //判斷副責任單位是否存在
                        if (SubResponsibleDeptId > 0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM BAS.Department a
                                WHERE a.DepartmentId = @SubResponsibleDeptId";
                            dynamicParameters.Add("SubResponsibleDeptId", SubResponsibleDeptId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【副責任單位】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //判斷副責任者是否存在
                        if (SubResponsibleUserId > 0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @SubResponsibleUserId
                                AND a.Status = 'A'";
                            dynamicParameters.Add("SubResponsibleUserId", SubResponsibleUserId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【副責任者】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //判斷編程者是否存在
                        if (ProgrammerUserId > 0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ProgrammerUserId
                                AND a.Status = 'A'";
                            dynamicParameters.Add("ProgrammerUserId", ProgrammerUserId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【編程者】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //判斷合致對象是否存在
                        if (ConformUserId > 0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ConformUserId
                                AND a.Status = 'A'";
                            dynamicParameters.Add("ConformUserId", ConformUserId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【合致對象】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //判斷對策代碼是否存在
                        if (RepairCauseUserId >= 0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                    FROM QMS.RepairCause a
                                    WHERE a.CauseId = @RepairCauseId
                                    AND a.Status = 'A'";
                            dynamicParameters.Add("RepairCauseId", RepairCauseId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【對策代碼】不存在，請重新輸入!");
                        }
                        else
                        {
                            RepairCauseId = nullData;
                        }
                        #endregion

                        #region //判斷編寫對策者是否存在
                        if (RepairCauseUserId >= 0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User] a
                                    WHERE a.UserId = @RepairCauseUserId
                                    AND a.Status = 'A'";
                            dynamicParameters.Add("RepairCauseUserId", RepairCauseUserId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【編寫對策者】不存在，請重新輸入!");
                        }
                        else
                        {
                            RepairCauseUserId = nullData;
                        }
                        #endregion

                        #endregion

                        #region //判斷品異目前處於得階段
                        if(RepairCauseUserId > 0)
                        {
                            AqBarcodeStatus = "2";
                        }
                        #endregion

                        #endregion

                        #region //品異單單身 - 異常條碼建立
                        if (QcType != "IQC")
                        {
                            foreach (var BarcodeId in BarcodeIds.Split(','))
                            {
                                #region //判斷條碼現在製程是否在異常製程
                                sql = @"SELECT a.BarcodeNo,a.BarcodeStatus,a.CurrentProdStatus
                                    FROM MES.Barcode a
                                    WHERE a.BarcodeId = @BarcodeId";
                                dynamicParameters.Add("BarcodeId", BarcodeId);
                                if (MoProcessId > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoProcessId", @" AND a.CurrentMoProcessId = @MoProcessId", MoProcessId);
                                var result1 = sqlConnection.Query(sql, dynamicParameters);
                                if (result1.Count() <= 0) throw new SystemException("【條碼Id:" + BarcodeId + "】不存在，請重新輸入!");
                                foreach (var item in result1)
                                {
                                    if (item.BarcodeStatus == "1" || item.BarcodeStatus == "0" || (item.BarcodeStatus == "7" && item.CurrentProdStatus == "NK"))
                                    {
                                        //可以繼續進行操作
                                    }
                                    else
                                    {
                                        throw new SystemException("【條碼Id:" + BarcodeId + "】非再製狀態或是加工完成狀態，不可以操作!");
                                    }
                                }
                                #endregion

                                #region //判斷條碼目前是否在品異單
                                sql = @"SELECT Top 1 a.AqBarcodeId ,a.JudgeStatus 
                                    FROM QMS.AqBarcode a
                                    WHERE a.BarcodeId =@BarcodeId
                                    Order By a.LastModifiedDate DESC
                                    ";
                                dynamicParameters.Add("BarcodeId", BarcodeId);
                                result1 = sqlConnection.Query(sql, dynamicParameters);
                                if (result1.Count() > 0)
                                {
                                    //判斷品異單判斷結果,如果有值且不是S代表已經判定完成,如果條碼有異常可以開立新的品異單
                                    string JudgeStatus = "";
                                    foreach (var item in result1)
                                    {
                                        JudgeStatus = item.JudgeStatus;
                                    }
                                    if (JudgeStatus.Length <= 0)
                                    {
                                        throw new SystemException("該條碼目前在品異單判定中，不可以輸入");
                                    }
                                    else if (JudgeStatus == "S")
                                    {
                                        //throw new SystemException("該條碼目前判定不良品，不可以輸入");
                                    }
                                }
                                #endregion

                                #region //更改異常條碼目前狀態
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.Barcode SET
                                    CurrentProdStatus = 'F',
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE BarcodeId = @BarcodeId";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      LastModifiedBy,
                                      LastModifiedDate,
                                      BarcodeId
                                  });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region 品異單身建立
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO QMS.AqBarcode (AbnormalqualityId, BarcodeId, DefectCauseId, DefectCauseDesc, ConformUserId, SupplierId,
                                    ResponsibleDeptId, ResponsibleUserId, SubResponsibleDeptId, SubResponsibleUserId, ProgrammerUserId, 
                                    RepairCauseId, RepairCauseDesc, RepairCauseUserId, AqBarcodeStatus, ChangeJudgeFlag,
                                    CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.AqBarcodeId
                                    VALUES (@AbnormalqualityId, @BarcodeId, @DefectCauseId, @DefectCauseDesc, @ConformUserId, @SupplierId,
                                    @ResponsibleDeptId, @ResponsibleUserId, @SubResponsibleDeptId, @SubResponsibleUserId, @ProgrammerUserId,
                                    @RepairCauseId, @RepairCauseDesc, @RepairCauseUserId, @AqBarcodeStatus, @ChangeJudgeFlag,
                                    @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        AbnormalqualityId,
                                        BarcodeId,
                                        DefectCauseId,
                                        DefectCauseDesc,
                                        ConformUserId = ConformUserId > 0 ? ConformUserId : nullData,
                                        SupplierId = SupplierId > 0 ? SupplierId : nullData,
                                        ResponsibleDeptId,
                                        ResponsibleUserId,
                                        SubResponsibleDeptId = SubResponsibleDeptId > 0 ? SubResponsibleDeptId : nullData,
                                        SubResponsibleUserId = SubResponsibleUserId > 0 ? SubResponsibleUserId : nullData,
                                        ProgrammerUserId = ProgrammerUserId > 0 ? ProgrammerUserId : nullData,
                                        AqBarcodeStatus,
                                        ChangeJudgeFlag,
                                        CreateDate,
                                        RepairCauseId,
                                        RepairCauseDesc,
                                        RepairCauseUserId,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult2.Count();
                                #endregion
                            }
                        }
                        else
                        {
                            #region 判斷進貨單是否可以開立
                            dynamicParameters = new DynamicParameters();
                                sql = @"SELECT b.AqBarcodeStatus
                                FROM QMS.Abnormalquality a
                                INNER JOIN QMS.AqBarcode b on a.AbnormalqualityId = b.AbnormalqualityId
                                WHERE a.GrDetailId = @GrDetailId
                                AND b.JudgeConfirm != 'Y'";
                            dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                            dynamicParameters.Add("GrDetailId", GrDetailId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("該進貨單目前尚存品異單據未完成判定!!不能開立新品異單");
                            #endregion


                            #region 品異單身建立
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.AqBarcode (AbnormalqualityId, BarcodeId, DefectCauseId, DefectCauseDesc, ConformUserId, SupplierId,
                                    ResponsibleDeptId, ResponsibleUserId, SubResponsibleDeptId, SubResponsibleUserId, ProgrammerUserId, 
                                    RepairCauseId, RepairCauseDesc, RepairCauseUserId, AqBarcodeStatus, ChangeJudgeFlag,
                                    CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.AqBarcodeId
                                    VALUES (@AbnormalqualityId, @BarcodeId, @DefectCauseId, @DefectCauseDesc, @ConformUserId, @SupplierId,
                                    @ResponsibleDeptId, @ResponsibleUserId, @SubResponsibleDeptId, @SubResponsibleUserId, @ProgrammerUserId,
                                    @RepairCauseId, @RepairCauseDesc, @RepairCauseUserId, @AqBarcodeStatus, @ChangeJudgeFlag,
                                    @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    AbnormalqualityId,
                                    BarcodeId = (int?)null,
                                    DefectCauseId,
                                    DefectCauseDesc,
                                    ConformUserId = ConformUserId > 0 ? ConformUserId : nullData,
                                    SupplierId = SupplierId > 0 ? SupplierId : nullData,
                                    ResponsibleDeptId,
                                    ResponsibleUserId,
                                    SubResponsibleDeptId = SubResponsibleDeptId > 0 ? SubResponsibleDeptId : nullData,
                                    SubResponsibleUserId = SubResponsibleUserId > 0 ? SubResponsibleUserId : nullData,
                                    ProgrammerUserId = ProgrammerUserId > 0 ? ProgrammerUserId : nullData,
                                    AqBarcodeStatus,
                                    ChangeJudgeFlag,
                                    CreateDate,
                                    RepairCauseId,
                                    RepairCauseDesc,
                                    RepairCauseUserId,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult2.Count();
                            #endregion

                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "新增成功",
                        });
                        #endregion

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

        #region //AddAbnormalqualityPadProject-- 新增品異單多項目 - 平板版本(只有回報) -- Shintokuro 2023-02-16
        public string AddAbnormalqualityPadProject(string AbnormalqualityData, string AbnormalProjectList)
        {
            try
            {
                if (AbnormalqualityData.Length <= 0) throw new SystemException("【品異單資料】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int AbnormalqualityId = 0;
                        int rowsAffected = 0;
                        int? MoId = null;
                        int? GrDetailId = null;
                        int? MoProcessId = null;
                        int? BarcodeId = null;
                        int? QcRecordId = null;
                        int? QcBarcodeId = null;
                        int? nullData = null;
                        string QcType = "";
                        var ConformUserId = nullData;
                        var SubResponsibleDeptId = nullData;
                        var SubResponsibleUserId = nullData;
                        var ProgrammerUserId = nullData;
                        var ResponsibleSupervisorId = nullData;
                        string DocDate = DateTime.Now.ToString("yyyyMM");


                        var AbnormalqualityJson = JObject.Parse(AbnormalqualityData);

                        int ResponsibleUserId = -1;
                        #region //品異單單頭建立
                        foreach (var item in AbnormalqualityJson["data"])
                        {
                            QcType = Convert.ToString(item["QcType"]);
                            ResponsibleUserId = Convert.ToInt32(item["ResponsibleUserId"]);
                            if (QcType == "IPQC" || QcType == "NON")
                            {
                                MoProcessId = Convert.ToInt32(item["MoProcessId"]);
                            }
                            else
                            {
                                MoProcessId = null;
                            }
                            if (Convert.ToInt32(item["ResponsibleDeptId"]) <= 0) throw new SystemException("【責任單位】不能為空!");
                            if (Convert.ToInt32(item["ResponsibleUserId"]) <= 0) throw new SystemException("【責任者】不能為空!");
                            if (QcType != "IQC")
                            {
                                if (Convert.ToInt32(item["MoId"]) <= 0) throw new SystemException("【製令】不能為空!");
                                MoId = Convert.ToInt32(item["MoId"]);
                                if (QcType != "PVTQC")
                                {
                                    if (Convert.ToInt32(item["BarcodeId"]) <= 0) throw new SystemException("【條碼】不能為空!");

                                    if (QcType != "NON")
                                    {
                                        if (Convert.ToInt32(item["QcBarcodeId"]) <= 0) throw new SystemException("【檢驗紀錄編號】不能為空!");
                                    }


                                    #region //判斷條碼是否有進品異
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.AqBarcodeId,a.ProcessStatus,a.JudgeStatus
                                        FROM QMS.AqBarcode a
								        WHERE BarcodeId = @BarcodeId";
                                    dynamicParameters.Add("BarcodeId", Convert.ToInt32(item["BarcodeId"]));
                                    var resultJudgeStatus = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultJudgeStatus.Count() > 0)
                                    {
                                        string JudgeStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).JudgeStatus;
                                        string ProcessStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ProcessStatus;
                                        int AqBarcodeId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).AqBarcodeId;
                                        if (JudgeStatus == "RW")
                                        {
                                            if (ProcessStatus == "I")
                                            {
                                                #region //資料更新
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE QMS.AqBarcode SET
                                                    ProcessStatus = @ProcessStatus,
                                                    LastModifiedBy = @LastModifiedBy,
                                                    LastModifiedDate = @LastModifiedDate
                                                    WHERE AqBarcodeId = @AqBarcodeId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      ProcessStatus = "V",
                                                      ConformUserId,
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      AqBarcodeId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                                #endregion
                                            }
                                        }
                                    }
                                    #endregion
                                }
                            }
                            else
                            {
                                if (Convert.ToInt32(item["GrDetailId"]) <= 0) throw new SystemException("【銷貨單】不能為空!");
                                GrDetailId = Convert.ToInt32(item["GrDetailId"]);
                                #region 判斷進貨單是否可以開立
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT b.AqBarcodeStatus
                                        FROM QMS.Abnormalquality a
                                        INNER JOIN QMS.AqBarcode b on a.AbnormalqualityId = b.AbnormalqualityId
                                        WHERE a.GrDetailId = @GrDetailId
                                        AND b.JudgeConfirm != 'Y'";
                                dynamicParameters.Add("GrDetailId", GrDetailId);
                                var result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0) throw new SystemException("該進貨單目前尚存品異單據未完成判定!!不能開立新品異單");
                                #endregion

                            }
                        }

                        #region //單號自動取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(AbnormalqualityNo))), '000'), 3)) + 1 CurrentNum
                                FROM QMS.Abnormalquality
								WHERE AbnormalqualityNo NOT LIKE '%[A-Za-z]%'";
                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                        string AbnormalqualityNo = DocDate + string.Format("{0:000}", currentNum);
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.Abnormalquality (CompanyId, MoId, GrDetailId, MoProcessId, AbnormalqualityNo, AbnormalqualityStatus, DocDate, QcType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.AbnormalqualityId
                                VALUES (@CompanyId, @MoId, @GrDetailId, @MoProcessId, @AbnormalqualityNo, @AbnormalqualityStatus, @DocDate, @QcType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                MoId,
                                GrDetailId,
                                MoProcessId,
                                AbnormalqualityNo,
                                AbnormalqualityStatus = "F",
                                DocDate = DateTime.Now.ToString("yyyyMMdd"),
                                QcType,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = ResponsibleUserId,
                                LastModifiedBy = ResponsibleUserId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        rowsAffected += insertResult.Count();

                        //取出單頭Id
                        foreach (var item in insertResult)
                        {
                            AbnormalqualityId = item.AbnormalqualityId;
                        }
                        #endregion

                        #region //品異單單身 - 異常條碼建立
                        foreach (var item in AbnormalqualityJson["data"])
                        {

                            if (QcType != "PVTQC" && QcType != "IQC")
                            {

                                if (QcType == "IPQC")
                                {
                                    #region //判斷條碼是否存在
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.BarcodeNo,a.BarcodeStatus
                                            FROM MES.Barcode a
                                            INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
		                                    INNER JOIN MES.BarcodeProcess c ON b.MoProcessId = c.MoProcessId and a.BarcodeId =c.BarcodeId
                                            WHERE a.BarcodeId = @BarcodeId
                                            --AND a.BarcodeStatus = '1' 
                                            AND　c.FinishDate is not null";
                                    dynamicParameters.Add("BarcodeId", Convert.ToInt32(item["BarcodeId"]));
                                    var result1 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result1.Count() <= 0) throw new SystemException("【條碼Id:" + Convert.ToInt32(item["BarcodeId"]) + "】不存在，請重新輸入!");
                                    #endregion
                                }
                                else if (QcType == "OQC")
                                {
                                    #region //判斷條碼是否存在
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.BarcodeNo,a.BarcodeStatus
                                            FROM MES.Barcode a
                                            INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
		                                    INNER JOIN MES.BarcodeProcess c ON b.MoProcessId = c.MoProcessId and a.BarcodeId =c.BarcodeId
                                            WHERE a.BarcodeId = @BarcodeId
                                            AND　c.FinishDate is not null";
                                    dynamicParameters.Add("BarcodeId", Convert.ToInt32(item["BarcodeId"]));
                                    var result1 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result1.Count() <= 0) throw new SystemException("【條碼Id:" + Convert.ToInt32(item["BarcodeId"]) + "】不存在，請重新輸入!");
                                    #endregion
                                }


                                #region //判斷條碼目前是否在品異單
                                sql = @"SELECT Top 1 a.AqBarcodeId ,a.JudgeStatus 
                                        FROM QMS.AqBarcode a
                                        WHERE a.BarcodeId =@BarcodeId
                                        Order By a.LastModifiedDate DESC
                                        ";
                                dynamicParameters.Add("BarcodeId", Convert.ToInt32(item["BarcodeId"]));
                                var result2 = sqlConnection.Query(sql, dynamicParameters);
                                if (result2.Count() > 0)
                                {
                                    //判斷品異單判斷結果,如果有值且不是S代表已經判定完成,如果條碼有異常可以開立新的品異單
                                    string JudgeStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).JudgeStatus;
                                    if (JudgeStatus == null)
                                    {
                                        throw new SystemException("【條碼Id:" + Convert.ToInt32(item["BarcodeId"]) + "】目前在品異單判定中，不可以開立品異單");
                                    }
                                    else if (JudgeStatus == "S")
                                    {
                                        throw new SystemException("【條碼Id:" + Convert.ToInt32(item["BarcodeId"]) + "】目前判定不良品，不可以開立品異單");
                                    }
                                }
                                #endregion
                            }

                            #region //判斷資料是否存在
                            #region //資料 - NOT NULL

                            #region //判斷檢驗紀錄編號是否存在
                            if (QcType != "PVTQC" && QcType != "IQC")
                            {
                                if (QcType != "NON")
                                {
                                    sql = @"SELECT TOP 1 1
                                        FROM MES.QcBarcode a
                                        WHERE a.QcBarcodeId = @QcBarcodeId";
                                    dynamicParameters.Add("QcBarcodeId", Convert.ToInt32(item["QcBarcodeId"]));
                                    var resultQcBarcodeId = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultQcBarcodeId.Count() <= 0) throw new SystemException("【檢驗紀錄編號:" + item["QcRecordId"].ToString() + "】不存在，請重新輸入!");
                                }
                            }
                            #endregion

                            #region //判斷責任單位是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b on a.DepartmentId = b.DepartmentId
                                WHERE b.DepartmentId = @ResponsibleDeptId";
                            dynamicParameters.Add("ResponsibleDeptId", Convert.ToInt32(item["ResponsibleDeptId"]));

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【責任單位】不存在，請重新輸入!");
                            #endregion

                            #region //判斷責任者是否存在
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ResponsibleUserId
                                AND a.Status = 'A'";
                            dynamicParameters.Add("ResponsibleUserId", Convert.ToInt32(item["ResponsibleUserId"]));

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【責任者】不存在，請重新輸入!");
                            #endregion
                            #endregion

                            #region //資料 - NULL
                            #region //判斷合致對象是否存在
                            if (item["ConformUserId"].ToString() != "")
                            {
                                dynamicParameters = new DynamicParameters();

                                sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ConformUserId
                                AND a.Status = 'A'";
                                dynamicParameters.Add("ConformUserId", Convert.ToInt32(item["ConformUserId"]));

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【合致對象】不存在，請重新輸入!");
                                ConformUserId = Convert.ToInt32(item["ConformUserId"]);
                            }
                            #endregion

                            #region //判斷副責任單位是否存在
                            if (item["SubResponsibleDeptId"].ToString() != "")
                            {
                                dynamicParameters = new DynamicParameters();

                                sql = @"SELECT TOP 1 1
                                FROM BAS.Department a
                                WHERE a.DepartmentId = @SubResponsibleDeptId";
                                dynamicParameters.Add("SubResponsibleDeptId", Convert.ToInt32(item["SubResponsibleDeptId"]));

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【副責任單位】不存在，請重新輸入!");
                                SubResponsibleDeptId = Convert.ToInt32(item["SubResponsibleDeptId"]);
                            }
                            #endregion

                            #region //判斷副責任者是否存在
                            if (item["SubResponsibleUserId"].ToString() != "")
                            {
                                dynamicParameters = new DynamicParameters();

                                sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @SubResponsibleUserId
                                AND a.Status = 'A'";
                                dynamicParameters.Add("SubResponsibleUserId", Convert.ToInt32(item["SubResponsibleUserId"]));

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【副責任者】不存在，請重新輸入!");

                                SubResponsibleUserId = Convert.ToInt32(item["SubResponsibleUserId"]);

                            }
                            #endregion

                            #region //判斷編程者是否存在
                            if (item["ProgrammerUserId"].ToString() != "")
                            {
                                dynamicParameters = new DynamicParameters();

                                sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ProgrammerUserId
                                AND a.Status = 'A'";
                                dynamicParameters.Add("ProgrammerUserId", Convert.ToInt32(item["ProgrammerUserId"]));

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【編程者】不存在，請重新輸入!");

                                ProgrammerUserId = Convert.ToInt32(item["ProgrammerUserId"]);

                            }
                            #endregion
                            #endregion
                            #endregion

                            

                            if (QcType == "PVTQC")
                            {
                                BarcodeId = nullData;
                                QcBarcodeId = nullData;

                                

                            }
                            else if(QcType == "IQC")
                            {
                                dynamicParameters = new DynamicParameters();

                                sql = @"SELECT TOP 1 1 
                                        FROM MES.QcRecord a
                                        WHERE a.QcRecordId = @QcRecordId";
                                dynamicParameters.Add("QcRecordId", Convert.ToInt32(item["QcRecordId"]));

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【檢驗紀錄編號】不存在，請重新輸入!");
                                QcRecordId = Convert.ToInt32(item["QcRecordId"]);
                            }
                            else
                            {
                                #region //更新異常條碼目前狀態
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.Barcode SET
                                        CurrentProdStatus = 'F',
                                        LastModifiedBy = @LastModifiedBy,
                                        LastModifiedDate = @LastModifiedDate
                                        WHERE BarcodeId = @BarcodeId";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      LastModifiedBy,
                                      LastModifiedDate,
                                      BarcodeId = Convert.ToInt32(item["BarcodeId"])
                                  });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                BarcodeId = Convert.ToInt32(item["BarcodeId"]);
                                if (QcType == "NON")
                                {
                                    if (item["QcBarcodeId"].Count() == 0)
                                    {
                                        QcBarcodeId = nullData;
                                    }
                                }
                                else
                                {
                                    QcBarcodeId = Convert.ToInt32(item["QcBarcodeId"]);
                                }
                            }


                            #region //新增 - 品異單單身
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.AqBarcode (AbnormalqualityId, BarcodeId, QcRecordId, QcBarcodeId, DefectCauseId, DefectCauseDesc, ConformUserId, 
                                    ResponsibleDeptId, ResponsibleUserId, SubResponsibleDeptId, SubResponsibleUserId, ProgrammerUserId, 
                                    RepairCauseId, RepairCauseDesc, RepairCauseUserId,AqBarcodeStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.AqBarcodeId
                                    VALUES (@AbnormalqualityId, @BarcodeId, @QcRecordId, @QcBarcodeId, @DefectCauseId, @DefectCauseDesc, @ConformUserId, 
                                    @ResponsibleDeptId, @ResponsibleUserId, @SubResponsibleDeptId, @SubResponsibleUserId, @ProgrammerUserId, 
                                    @RepairCauseId, @RepairCauseDesc, @RepairCauseUserId, @AqBarcodeStatus, 
                                    @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    AbnormalqualityId,
                                    BarcodeId,
                                    QcRecordId,
                                    QcBarcodeId,
                                    DefectCauseId = QcType != "IQC" ? nullData : Convert.ToInt32(item["DefectCauseId"]),
                                    DefectCauseDesc = QcType != "IQC" ? (string)null : item["DefectCauseDesc"].ToString(),
                                    ConformUserId,
                                    ResponsibleDeptId = Convert.ToInt32(item["ResponsibleDeptId"]),
                                    ResponsibleUserId = Convert.ToInt32(item["ResponsibleUserId"]),
                                    SubResponsibleDeptId,
                                    SubResponsibleUserId,
                                    ProgrammerUserId,
                                    RepairCauseId = nullData,
                                    RepairCauseDesc = nullData,
                                    RepairCauseUserId = nullData,
                                    AqBarcodeStatus = 1,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult2.Count();

                            int AqBarcodeId = 0;
                            foreach (var item2 in insertResult2)
                            {
                                AqBarcodeId = item2.AqBarcodeId;
                            }
                            #endregion

                            if (QcType != "IQC")
                            {
                                if (AbnormalProjectList.Length <= 0) throw new SystemException("【量測異常資料】不能為空!");
                                var AbnormalProjectListJson = JObject.Parse(AbnormalProjectList);

                                foreach (var item2 in AbnormalProjectListJson["data"])
                                {
                                    if (Convert.ToInt32(item2["CauseId"]) <= 0) throw new SystemException("【不良代碼】不能為空!");
                                    if (item2["CauseDesc"].ToString().Length > 100) throw new SystemException("【不良原因】長度錯誤!");

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO QMS.AqQcItem (AqBarcodeId, QcItemId, DefectCauseId, DefectCauseDesc, RepairCauseId, RepairCauseDesc
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.AqQcItemId
                                        VALUES (@AqBarcodeId, @QcItemId, @DefectCauseId, @DefectCauseDesc, @RepairCauseId, @RepairCauseDesc
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            AqBarcodeId,
                                            QcItemId = Convert.ToInt32(item2["QcItemId"]),
                                            DefectCauseId = Convert.ToInt32(item2["CauseId"]),
                                            DefectCauseDesc = item2["CauseDesc"].ToString(),
                                            RepairCauseId = nullData,
                                            RepairCauseDesc = nullData,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy = ResponsibleUserId,
                                            LastModifiedBy = ResponsibleUserId
                                        });
                                    var insertAqQcItemResult = sqlConnection.Query(sql, dynamicParameters);
                                    rowsAffected += insertAqQcItemResult.Count();
                                }
                            }
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

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

        #region //AddAbnormalqualityPad-- 新增品異單 - 平板版本(到對策確認前) -- Shintokuro 2022-10-04
        public string AddAbnormalqualityPad(string AbnormalqualityData)
        {
            try
            {
                if (AbnormalqualityData.Length <= 0) throw new SystemException("【品異單資料】不能為空!");
               
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int AbnormalqualityId = 0;
                        int rowsAffected = 0;
                        int MoId = 0;
                        int? MoProcessId = 0;
                        int? BarcodeId = 0;
                        int? QcBarcodeId = 0;
                        int? nullData = null;
                        string QcType = "";
                        var ConformUserId = nullData;
                        var SubResponsibleDeptId = nullData;
                        var SubResponsibleUserId = nullData;
                        var ProgrammerUserId = nullData;
                        var RepairCauseId = nullData;
                        var RepairCauseUserId = nullData;
                        //var RepairCauseDesc = nullData;
                        var ResponsibleSupervisorId = nullData;
                        string DocDate = DateTime.Now.ToString("yyyyMM");
                        int ResponsibleUserId = -1;
                        #region //品異單單頭建立
                        var AbnormalqualityJson = JObject.Parse(AbnormalqualityData);

                        foreach (var item in AbnormalqualityJson["data"])
                        {
                            if (Convert.ToInt32(item["MoId"]) <= 0) throw new SystemException("【製令】不能為空!");
                            if (Convert.ToInt32(item["DefectCauseId"]) <= 0) throw new SystemException("【不良代碼】不能為空!");
                            if (Convert.ToInt32(item["ResponsibleDeptId"]) <= 0) throw new SystemException("【責任單位】不能為空!");
                            if (Convert.ToInt32(item["ResponsibleUserId"]) <= 0) throw new SystemException("【責任者】不能為空!");
                            //if(item["DefectCauseDesc"].ToString() != null)
                            //{
                            //    if (item["DefectCauseDesc"].ToString().Length > 100) throw new SystemException("【不良原因】長度錯誤!");
                            //}

                            if (Convert.ToString(item["QcType"]) != "PVTQC")
                            {
                                if (Convert.ToInt32(item["BarcodeId"]) <= 0) throw new SystemException("【條碼】不能為空!");

                                if (Convert.ToString(item["QcType"]) != "NON")
                                {
                                    if (Convert.ToInt32(item["QcBarcodeId"]) <= 0) throw new SystemException("【檢驗紀錄編號】不能為空!");
                                }


                                #region //判斷條碼是否有進品異
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.AqBarcodeId,a.ProcessStatus,a.JudgeStatus
                                        FROM QMS.AqBarcode a
								        WHERE BarcodeId = @BarcodeId";
                                dynamicParameters.Add("BarcodeId", Convert.ToInt32(item["BarcodeId"]));
                                var resultJudgeStatus = sqlConnection.Query(sql, dynamicParameters);
                                if (resultJudgeStatus.Count() > 0)
                                {
                                    string JudgeStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).JudgeStatus;
                                    string ProcessStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ProcessStatus;
                                    int AqBarcodeId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).AqBarcodeId;
                                    if (JudgeStatus == "RW")
                                    {
                                        if (ProcessStatus == "I")
                                        {
                                            #region //資料更新
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE QMS.AqBarcode SET
                                                    ProcessStatus = @ProcessStatus,
                                                    LastModifiedBy = @LastModifiedBy,
                                                    LastModifiedDate = @LastModifiedDate
                                                    WHERE AqBarcodeId = @AqBarcodeId";
                                            dynamicParameters.AddDynamicParams(
                                              new
                                              {
                                                  ProcessStatus = "V",
                                                  ConformUserId,
                                                  LastModifiedBy,
                                                  LastModifiedDate,
                                                  AqBarcodeId
                                              });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion
                                        }
                                    }
                                }
                                #endregion
                            }


                            MoId = Convert.ToInt32(item["MoId"]);
                            QcType = Convert.ToString(item["QcType"]);
                            if(QcType == "IPQC" || QcType == "NON")
                            {
                                MoProcessId = Convert.ToInt32(item["MoProcessId"]);
                            }
                            else
                            {
                                MoProcessId = null;
                            }
                            //DocDate = item["Docate"].ToString();
                            ResponsibleUserId = Convert.ToInt32(item["ResponsibleUserId"]);
                        }

        //                #region //判斷製令是否批量,批量只能做出貨檢
        //                dynamicParameters = new DynamicParameters();
        //                sql = @"SELECT a.LotStatus
        //                        FROM MES.MoSetting a
								//WHERE MoId = @MoId";
        //                dynamicParameters.Add("MoId", MoId);
        //                var result = sqlConnection.Query(sql, dynamicParameters);
        //                if(result.Count() <= 0) throw new SystemException("資料有問題,找不到該製令設定");
        //                string LotStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).LotStatus;
        //                if(LotStatus == "Y")
        //                {
        //                    if(QcType == "IPQC") throw new SystemException("該製令為批量生產,只能做出貨檢或是試產檢");
        //                }
        //                #endregion


                        #region //單號自動取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(AbnormalqualityNo))), '000'), 3)) + 1 CurrentNum
                                FROM QMS.Abnormalquality
								WHERE AbnormalqualityNo NOT LIKE '%[A-Za-z]%'";
                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                        string AbnormalqualityNo = DocDate + string.Format("{0:000}", currentNum);
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.Abnormalquality (CompanyId, MoId, MoProcessId, AbnormalqualityNo, AbnormalqualityStatus, DocDate, QcType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.AbnormalqualityId
                                VALUES (@CompanyId, @MoId, @MoProcessId, @AbnormalqualityNo, @AbnormalqualityStatus, @DocDate, @QcType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                MoId,
                                MoProcessId,
                                AbnormalqualityNo,
                                AbnormalqualityStatus = "F",
                                DocDate = DateTime.Now.ToString("yyyyMMdd"),
                                QcType,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = ResponsibleUserId,
                                LastModifiedBy = ResponsibleUserId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        rowsAffected += insertResult.Count();

                        //取出單頭Id
                        foreach (var item in insertResult)
                        {
                            AbnormalqualityId = item.AbnormalqualityId;
                        }
                        #endregion

                        #region //品異單單身 - 異常條碼建立
                        foreach (var item in AbnormalqualityJson["data"])
                        {
                            
                            if (QcType != "PVTQC")
                            {

                                if(QcType == "IPQC")
                                {
                                    #region //判斷條碼是否存在
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.BarcodeNo,a.BarcodeStatus
                                        FROM MES.Barcode a
                                        INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
		                                INNER JOIN MES.BarcodeProcess c ON b.MoProcessId = c.MoProcessId and a.BarcodeId =c.BarcodeId
                                        WHERE a.BarcodeId = @BarcodeId
                                        --AND a.BarcodeStatus = '1' 
                                        AND　c.FinishDate is not null";
                                    dynamicParameters.Add("BarcodeId", Convert.ToInt32(item["BarcodeId"]));
                                    var result1 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result1.Count() <= 0) throw new SystemException("【條碼Id:" + Convert.ToInt32(item["BarcodeId"]) + "】不存在，請重新輸入!");
                                    #endregion
                                }
                                else if(QcType == "OQC")
                                {
                                    #region //判斷條碼是否存在
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.BarcodeNo,a.BarcodeStatus
                                            FROM MES.Barcode a
                                            INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
		                                    INNER JOIN MES.BarcodeProcess c ON b.MoProcessId = c.MoProcessId and a.BarcodeId =c.BarcodeId
                                            WHERE a.BarcodeId = @BarcodeId
                                            AND　c.FinishDate is not null";
                                    dynamicParameters.Add("BarcodeId", Convert.ToInt32(item["BarcodeId"]));
                                    var result1 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result1.Count() <= 0) throw new SystemException("【條碼Id:" + Convert.ToInt32(item["BarcodeId"]) + "】不存在，請重新輸入!");
                                    #endregion
                                }


                                #region //判斷條碼目前是否在品異單
                                sql = @"SELECT Top 1 a.AqBarcodeId ,a.JudgeStatus 
                                    FROM QMS.AqBarcode a
                                    WHERE a.BarcodeId =@BarcodeId
                                    Order By a.LastModifiedDate DESC
                                ";
                                dynamicParameters.Add("BarcodeId", Convert.ToInt32(item["BarcodeId"]));
                                var result2 = sqlConnection.Query(sql, dynamicParameters);
                                if (result2.Count() > 0)
                                {
                                    //判斷品異單判斷結果,如果有值且不是S代表已經判定完成,如果條碼有異常可以開立新的品異單
                                    string JudgeStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).JudgeStatus;
                                    if (JudgeStatus == null)
                                    {
                                        throw new SystemException("【條碼Id:" + Convert.ToInt32(item["BarcodeId"]) + "】目前在品異單判定中，不可以開立品異單");
                                    }
                                    else if (JudgeStatus == "S")
                                    {
                                        throw new SystemException("【條碼Id:" + Convert.ToInt32(item["BarcodeId"]) + "】目前判定不良品，不可以開立品異單");
                                    }
                                }
                                #endregion
                            }

                            #region //判斷資料是否存在
                            #region //資料 - NOT NULL

                            


                            #region //判斷不良原因代碼是否存在
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 CauseDesc
                                FROM QMS.DefectCause a
                                WHERE a.CauseId = @DefectCauseId";
                            dynamicParameters.Add("DefectCauseId", Convert.ToInt32(item["DefectCauseId"]));

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【不良原因代碼】不存在，請重新輸入!");
                            string DefectCauseDesc = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CauseDesc;
                            #endregion

                            if (QcType != "PVTQC")
                            {
                                if(QcType != "NON")
                                {
                                    #region //判斷檢驗紀錄編號是否存在
                                    sql = @"SELECT TOP 1 1
                                        FROM MES.QcBarcode a
                                        WHERE a.QcBarcodeId = @QcBarcodeId";
                                    dynamicParameters.Add("QcBarcodeId", Convert.ToInt32(item["QcBarcodeId"]));
                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【檢驗紀錄編號:" + item["QcRecordId"].ToString() + "】不存在，請重新輸入!");
                                    #endregion
                                }
                            }

                            #region //判斷責任單位是否存在
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b on a.DepartmentId = b.DepartmentId
                                WHERE b.DepartmentId = @ResponsibleDeptId";
                            dynamicParameters.Add("ResponsibleDeptId", Convert.ToInt32(item["ResponsibleDeptId"]));

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【責任單位】不存在，請重新輸入!");
                            #endregion

                            #region //判斷責任者是否存在
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ResponsibleUserId
                                AND a.Status = 'A'";
                            dynamicParameters.Add("ResponsibleUserId", Convert.ToInt32(item["ResponsibleUserId"]));

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【責任者】不存在，請重新輸入!");
                            #endregion
                            #endregion

                            #region //資料 - NULL
                            #region //判斷合致對象是否存在
                            if (item["ConformUserId"].ToString() != "")
                            {
                                dynamicParameters = new DynamicParameters();

                                sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ConformUserId
                                AND a.Status = 'A'";
                                dynamicParameters.Add("ConformUserId", Convert.ToInt32(item["ConformUserId"]));

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【合致對象】不存在，請重新輸入!");
                                ConformUserId = Convert.ToInt32(item["ConformUserId"]);
                            }
                            #endregion

                            #region //判斷副責任單位是否存在
                            if (item["SubResponsibleDeptId"].ToString() != "")
                            {
                                dynamicParameters = new DynamicParameters();

                                sql = @"SELECT TOP 1 1
                                FROM BAS.Department a
                                WHERE a.DepartmentId = @SubResponsibleDeptId";
                                dynamicParameters.Add("SubResponsibleDeptId", Convert.ToInt32(item["SubResponsibleDeptId"]));

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【副責任單位】不存在，請重新輸入!");
                                SubResponsibleDeptId = Convert.ToInt32(item["SubResponsibleDeptId"]);
                            }
                            #endregion

                            #region //判斷副責任者是否存在
                            if (item["SubResponsibleUserId"].ToString() != "")
                            {
                                dynamicParameters = new DynamicParameters();

                                sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @SubResponsibleUserId
                                AND a.Status = 'A'";
                                dynamicParameters.Add("SubResponsibleUserId", Convert.ToInt32(item["SubResponsibleUserId"]));

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【副責任者】不存在，請重新輸入!");

                                SubResponsibleUserId = Convert.ToInt32(item["SubResponsibleUserId"]);
                               
                            }
                            #endregion

                            #region //判斷編程者是否存在
                            if (item["ProgrammerUserId"].ToString() != "")
                            {
                                dynamicParameters = new DynamicParameters();

                                sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ProgrammerUserId
                                AND a.Status = 'A'";
                                dynamicParameters.Add("ProgrammerUserId", Convert.ToInt32(item["ProgrammerUserId"]));

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【編程者】不存在，請重新輸入!");

                                ProgrammerUserId = Convert.ToInt32(item["ProgrammerUserId"]);
                                
                            }
                            #endregion

                            //#region //判斷對策代碼是否存在
                            //if (item["RepairCauseId"].ToString() != "")
                            //{
                            //    dynamicParameters = new DynamicParameters();

                            //    sql = @"SELECT TOP 1 1
                            //    FROM QMS.RepairCause a
                            //    WHERE a.CauseId = @RepairCauseId";
                            //    dynamicParameters.Add("RepairCauseId", Convert.ToInt32(item["RepairCauseId"]));

                            //    result = sqlConnection.Query(sql, dynamicParameters);
                            //    if (result.Count() <= 0) throw new SystemException("【對策代碼】不存在，請重新輸入!");

                            //    RepairCauseId = Convert.ToInt32(item["RepairCauseId"]);
                                
                            //}
                            //#endregion

                            //#region //判斷對策原因長度是否正確
                            //if (item["RepairCauseDesc"].ToString().Length > 0)
                            //{
                            //    if (item["RepairCauseDesc"].ToString().Length > 100) throw new SystemException("【對策原因】長度錯誤!");
                            //}
                            //#endregion

                            //#region //判斷編寫對策者是否存在
                            //if (item["RepairCauseUserId"].ToString() != "")
                            //{
                            //    dynamicParameters = new DynamicParameters();

                            //    sql = @"SELECT TOP 1 1
                            //    FROM BAS.[User] a
                            //    WHERE a.UserId = @RepairCauseUserId
                            //    AND a.Status = 'A'";
                            //    dynamicParameters.Add("RepairCauseUserId", Convert.ToInt32(item["RepairCauseUserId"]));

                            //    result = sqlConnection.Query(sql, dynamicParameters);
                            //    if (result.Count() <= 0) throw new SystemException("【編寫對策者】不存在，請重新輸入!");

                            //    RepairCauseUserId = Convert.ToInt32(item["RepairCauseUserId"]);

                            //}
                            //#endregion

                            #endregion
                            #endregion

                            #region //判斷異常條碼目前處在哪個階段
                            int AqBarcodeStatus = 0;
                            if (Convert.ToInt32(item["DefectCauseId"]) > 0)
                            {
                                AqBarcodeStatus = 1;
                            }
                            else
                            {
                                throw new SystemException("【缺少不良原因代碼】，不能回報");
                            }
                            #endregion

                            if(QcType != "PVTQC")
                            {
                                #region //更新異常條碼目前狀態
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.Barcode SET
                                        CurrentProdStatus = 'F',
                                        LastModifiedBy = @LastModifiedBy,
                                        LastModifiedDate = @LastModifiedDate
                                        WHERE BarcodeId = @BarcodeId";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      LastModifiedBy,
                                      LastModifiedDate,
                                      BarcodeId = Convert.ToInt32(item["BarcodeId"])
                                  });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                BarcodeId = Convert.ToInt32(item["BarcodeId"]);
                                if (QcType == "NON")
                                {
                                    if (item["QcBarcodeId"].Count() == 0)
                                    {
                                        QcBarcodeId = nullData;
                                    }
                                }
                                else
                                {
                                    QcBarcodeId = Convert.ToInt32(item["QcBarcodeId"]);
                                }

                            }
                            else
                            {
                                BarcodeId = nullData;
                                QcBarcodeId = nullData;
                            }



                            #region //新增 - 品異單單身
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.AqBarcode (AbnormalqualityId, BarcodeId, QcBarcodeId, DefectCauseId, DefectCauseDesc, ConformUserId, 
                                    ResponsibleDeptId, ResponsibleUserId, SubResponsibleDeptId, SubResponsibleUserId, ProgrammerUserId, 
                                    RepairCauseId, RepairCauseDesc, RepairCauseUserId,AqBarcodeStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.AqBarcodeId
                                    VALUES (@AbnormalqualityId, @BarcodeId, @QcBarcodeId, @DefectCauseId, @DefectCauseDesc, @ConformUserId, 
                                    @ResponsibleDeptId, @ResponsibleUserId, @SubResponsibleDeptId, @SubResponsibleUserId, @ProgrammerUserId, 
                                    @RepairCauseId, @RepairCauseDesc, @RepairCauseUserId, @AqBarcodeStatus, 
                                    @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    AbnormalqualityId,
                                    BarcodeId,
                                    QcBarcodeId,
                                    DefectCauseId = Convert.ToInt32(item["DefectCauseId"]),
                                    DefectCauseDesc,
                                    ConformUserId,
                                    ResponsibleDeptId = Convert.ToInt32(item["ResponsibleDeptId"]),
                                    ResponsibleUserId = Convert.ToInt32(item["ResponsibleUserId"]),
                                    SubResponsibleDeptId,
                                    SubResponsibleUserId,
                                    ProgrammerUserId,
                                    RepairCauseId,
                                    RepairCauseDesc = nullData,
                                    RepairCauseUserId,
                                    AqBarcodeStatus,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult2.Count();
                            #endregion

                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

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

        #region //AddAqFile-- 新增品異條碼原因佐證檔案 -- Shintokuro 2022-10-14
        public string AddAqFile(int AqBarcodeId, string FileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //品異條碼原因佐證檔案建立

                        foreach (var item in FileId.Split(','))
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.AqFile (AqBarcodeId, AqFileStatus, FileId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.AqFileId
                                    VALUES (@AqBarcodeId, @AqFileStatus, @FileId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    AqBarcodeId,
                                    AqFileStatus = 1,
                                    FileId = item,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                        });
                        #endregion

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

        #region //Update
        #region //UpdateAqReturnDetail -- 更新品異單異常回報資料 -- Shintokuro 2022.10.12
        public string UpdateAqReturnDetail(int AbnormalqualityId, int AqBarcodeId, int DefectCauseId, string DefectCauseDesc, int ResponsibleDeptId
            , int ResponsibleUserId, int SubResponsibleDeptId, int SubResponsibleUserId, int ProgrammerUserId ,int ConformUserId, int? RepairCauseId, string RepairCauseDesc, int? RepairCauseUserId
            , int SupplierId, string ChangeJudgeFlag)
        {
            try
            {
                if (DefectCauseId <= 0) throw new SystemException("不良代碼不能為空,請重新確認");
                if (ResponsibleDeptId <= 0) throw new SystemException("對責任者為空,請重新確認");
                if (ResponsibleUserId <= 0) throw new SystemException("對責任部門不能為空,請重新確認");
                if (DefectCauseDesc.ToString().Length > 100) throw new SystemException("【不良原因】長度錯誤!");
                if (RepairCauseDesc.ToString().Length > 100) throw new SystemException("【對策原因】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        int? nullData = null;
                        string JudgeConfirm = "";
                        int? ResponsibleSupervisorId = -1;//責任主管
                        int AqBarcodeStatus = -1;//當前-異常條碼階段
                        int AqBarcodeStatusBase = -1;//資料庫-異常條碼階段

                        #region //判斷異常條碼資料是否正確+判斷是否相同人修改
                        sql = @"SELECT ResponsibleUserId, JudgeConfirm, ResponsibleSupervisorId, AqBarcodeStatus
                                FROM QMS.AqBarcode
                                WHERE AbnormalqualityId = @AbnormalqualityId
                                AND AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("資料不存在,請重新確認");
                        foreach(var item  in result)
                        {
                            JudgeConfirm = item.JudgeConfirm;
                            ResponsibleSupervisorId = item.ResponsibleSupervisorId;
                            AqBarcodeStatusBase =Convert.ToInt32(item.AqBarcodeStatus);
                        }
                        //int ResponsibleUserIdOriginal = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ResponsibleUserId;
                        //if(ResponsibleUserIdOriginal != CreateBy) throw new SystemException("請不要修改他人資料!!!");
                        if(AqBarcodeStatusBase == 3) throw new SystemException("該異常條碼已經完成對策確認,不可再更動");
                        if(AqBarcodeStatusBase == 5) throw new SystemException("該異常條碼已經完成會簽,不可再更動");
                        #endregion

                        #region //判斷不良原因代碼是否存在
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectCause a
                                WHERE a.CauseId = @DefectCauseId";
                        dynamicParameters.Add("DefectCauseId", DefectCauseId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【不良原因代碼】不存在，請重新輸入!");
                        #endregion

                        #region //判斷責任單位是否存在
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department a
                                WHERE a.DepartmentId = @ResponsibleDeptId";
                        dynamicParameters.Add("ResponsibleDeptId", ResponsibleDeptId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【責任單位】不存在，請重新輸入!");
                        #endregion

                        #region //判斷責任者是否存在
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ResponsibleUserId
                                AND a.Status = 'A'";
                        dynamicParameters.Add("ResponsibleUserId", ResponsibleUserId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【責任者】不存在，請重新輸入!");
                        #endregion

                        #region //判斷副責任單位是否存在
                        if (SubResponsibleDeptId >=0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM BAS.Department a
                                WHERE a.DepartmentId = @SubResponsibleDeptId";
                            dynamicParameters.Add("SubResponsibleDeptId", SubResponsibleDeptId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【副責任單位】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //判斷副責任者是否存在
                        if (SubResponsibleUserId >= 0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @SubResponsibleUserId
                                AND a.Status = 'A'";
                            dynamicParameters.Add("SubResponsibleUserId", SubResponsibleUserId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【副責任者】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //判斷編程者是否存在
                        if (ProgrammerUserId >= 0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ProgrammerUserId
                                AND a.Status = 'A'";
                            dynamicParameters.Add("ProgrammerUserId", ProgrammerUserId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【編程者】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //判斷合致對象是否存在
                        if (ConformUserId >= 0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ConformUserId
                                AND a.Status = 'A'";
                            dynamicParameters.Add("ConformUserId", ConformUserId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【合致對象】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //判斷編寫對策者是否存在
                        if (RepairCauseUserId >= 0)
                        {
                            if (RepairCauseId <= 0) throw new SystemException("【對策代碼】不能為空,請重新確認");

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User] a
                                    WHERE a.UserId = @RepairCauseUserId
                                    AND a.Status = 'A'";
                            dynamicParameters.Add("RepairCauseUserId", RepairCauseUserId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【編寫對策者】不存在，請重新輸入!");

                            dynamicParameters = new DynamicParameters();

                            #region //判斷對策代碼是否存在
                            sql = @"SELECT TOP 1 1
                                    FROM QMS.RepairCause a
                                    WHERE a.CauseId = @RepairCauseId
                                    AND a.Status = 'A'";
                            dynamicParameters.Add("RepairCauseId", RepairCauseId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【對策代碼】不存在，請重新輸入!");
                            #endregion
                        }
                        else
                        {
                            RepairCauseUserId = nullData;
                        }
                        #endregion

                        #region //異常條碼目前階段判定
                        if(JudgeConfirm == "Y") //判定確認
                        {
                            if(RepairCauseId > 0 && ResponsibleSupervisorId > 0) //對策確認完成
                            {
                                AqBarcodeStatus = 3;
                            }
                            else if(RepairCauseUserId > 0) // 對策完成
                            {
                                AqBarcodeStatus = 2;
                            }
                            else //只做判定
                            {
                                AqBarcodeStatus = 4;
                            }
                        }
                        else
                        {
                            if (RepairCauseId > 0 && ResponsibleSupervisorId > 0) //對策確認完成
                            {
                                AqBarcodeStatus = 3;
                            }
                            else if (RepairCauseUserId > 0)//對策完成
                            {
                                AqBarcodeStatus = 2;
                            }
                            else //只有回報
                            {
                                AqBarcodeStatus = 1;
                            }
                        }
                        #endregion


                        #region //資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.AqBarcode SET
                                    DefectCauseId = @DefectCauseId,
                                    DefectCauseDesc = @DefectCauseDesc,
                                    ResponsibleDeptId = @ResponsibleDeptId,
                                    ResponsibleUserId = @ResponsibleUserId,
                                    SubResponsibleDeptId = @SubResponsibleDeptId,
                                    SubResponsibleUserId = @SubResponsibleUserId,
                                    ProgrammerUserId = @ProgrammerUserId,
                                    ConformUserId = @ConformUserId,
                                    SupplierId = @SupplierId,
                                    RepairCauseId = @RepairCauseId,
                                    RepairCauseDesc = @RepairCauseDesc,
                                    RepairCauseUserId = @RepairCauseUserId,
                                    AqBarcodeStatus = @AqBarcodeStatus,
                                    ChangeJudgeFlag = @ChangeJudgeFlag,
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE AbnormalqualityId = @AbnormalqualityId
                                    AND AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              DefectCauseId,
                              DefectCauseDesc,
                              ResponsibleDeptId,
                              ResponsibleUserId,
                              SubResponsibleDeptId = SubResponsibleUserId > 0 ? SubResponsibleDeptId : nullData,
                              SubResponsibleUserId = SubResponsibleUserId > 0 ? SubResponsibleUserId : nullData,
                              ProgrammerUserId = ProgrammerUserId > 0 ? ProgrammerUserId : nullData,
                              ConformUserId = ConformUserId > 0 ? ConformUserId : nullData,
                              SupplierId = SupplierId > 0 ? SupplierId : nullData,
                                  //RepairCauseId,
                                  //RepairCauseDesc,
                                  RepairCauseId = RepairCauseId > 0 ? RepairCauseId : nullData,
                              RepairCauseDesc = RepairCauseId > 0 ? RepairCauseDesc : nullData.ToString(),
                              RepairCauseUserId,
                              AqBarcodeStatus,
                              ChangeJudgeFlag,
                              LastModifiedBy,
                              LastModifiedDate,
                              AbnormalqualityId,
                              AqBarcodeId
                          });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
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

        #region //UpdateAqCountermeasureDetail -- 更新品異單異常對策資料 -- Shintokuro 2022.10.12
        public string UpdateAqCountermeasureDetail(int AbnormalqualityId, int AqBarcodeId, int RepairCauseId, string RepairCauseDesc, int RepairCauseUserId, string RepairCauseList
            ,int ResponsibleDeptId, int ResponsibleUserId)
        {
            try
            {
                if (AbnormalqualityId <= 0) throw new SystemException("品異單據不能為空,請重新確認");
                if (AqBarcodeId <= 0) throw new SystemException("不良條碼不能為空,請重新確認");
                if (ResponsibleDeptId <= 0) throw new SystemException("【責任單位】不能為空,請重新確認");
                if (ResponsibleUserId <= 0) throw new SystemException("【責任者】不能為空,請重新確認");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();
                        int rowsAffected = 0;
                        int? RepairCauseUserIdBase=-1;  //資料庫-編寫對策者
                        int? ResponsibleSupervisorId = -1;  //資料庫-責任主管
                        int? DefectCauseIdBase = -1;    //資料庫-不良代碼
                        int AqBarcodeStatus = -1;       //當前-異常條碼狀態
                        int AqBarcodeStatusBase = -1;   //資料庫-異常條碼狀態
                        string JudgeConfirm = "";       //資料庫-判定狀態

                        #region //判斷異常條碼資料是否正確+判斷是否相同人修改
                        sql = @"SELECT RepairCauseUserId,AqBarcodeStatus,DefectCauseId, AqBarcodeStatus, JudgeConfirm, ResponsibleSupervisorId
                                FROM QMS.AqBarcode
                                WHERE AbnormalqualityId = @AbnormalqualityId
                                AND AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("資料不存在,請重新確認");
                        foreach(var item in result)
                        {
                            RepairCauseUserIdBase = item.RepairCauseUserId; 
                            DefectCauseIdBase = item.DefectCauseId;
                            AqBarcodeStatusBase = Convert.ToInt32(item.AqBarcodeStatus);
                            JudgeConfirm = item.JudgeConfirm;
                        }
                        if (AqBarcodeStatusBase == 3) throw new SystemException("該異常條碼已經完成對策確認,不可再更動");
                        if (AqBarcodeStatusBase == 5) throw new SystemException("該異常條碼已經完成會簽,不可再更動");

                        //if (RepairCauseUserIdBase != null)@Lo01032
                        //{
                        //    if (RepairCauseUserIdBase != CreateBy) throw new SystemException("請不要修改他人資料!!!");
                        //}
                        #endregion

                        #region //判斷責任單位是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM  BAS.Department a
                                WHERE a.DepartmentId = @ResponsibleDeptId";
                        dynamicParameters.Add("ResponsibleDeptId", ResponsibleDeptId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【責任單位】不存在，請重新輸入!");
                        #endregion

                        #region //判斷責任者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ResponsibleUserId
                                AND a.Status = 'A'";
                        dynamicParameters.Add("ResponsibleUserId", ResponsibleUserId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【責任者】不存在，請重新輸入!");
                        #endregion

                        

                        if (DefectCauseIdBase != null) //單筆異常回報
                        {
                            if (RepairCauseId <= 0) throw new SystemException("對策代碼不能為空,請重新確認");
                            if (RepairCauseUserId <= 0) throw new SystemException("編寫對策者不能為空,請重新確認");

                            #region //判斷對策代碼是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM QMS.RepairCause a
                                    WHERE a.CauseId = @RepairCauseId";
                            dynamicParameters.Add("RepairCauseId", RepairCauseId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【對策代碼】不存在，請重新輸入!");
                            #endregion

                            #region //異常條碼目前階段判定
                            if (JudgeConfirm == "Y") //判定確認
                            {
                                if (RepairCauseId > 0 && ResponsibleSupervisorId > 0) //對策確認完成
                                {
                                    AqBarcodeStatus = 3;
                                }
                                else if (RepairCauseUserId > 0) // 對策完成
                                {
                                    AqBarcodeStatus = 2;
                                }
                                else //只做判定
                                {
                                    AqBarcodeStatus = 4;
                                }
                            }
                            else
                            {
                                if (RepairCauseId > 0 && ResponsibleSupervisorId > 0) //對策確認完成
                                {
                                    AqBarcodeStatus = 3;
                                }
                                else if (RepairCauseUserId > 0)//對策完成
                                {
                                    AqBarcodeStatus = 2;
                                }
                                else //只有回報
                                {
                                    AqBarcodeStatus = 1;
                                }
                            }
                            #endregion

                            #region //更新-異常條碼的對策資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.AqBarcode SET
                                    RepairCauseId = @RepairCauseId,
                                    RepairCauseDesc = @RepairCauseDesc,
                                    RepairCauseUserId = @RepairCauseUserId,
                                    ResponsibleDeptId = @ResponsibleDeptId,
                                    ResponsibleUserId = @ResponsibleUserId,
                                    AqBarcodeStatus = @AqBarcodeStatus,
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE AbnormalqualityId = @AbnormalqualityId
                                    AND AqBarcodeId = @AqBarcodeId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  RepairCauseId,
                                  RepairCauseDesc,
                                  RepairCauseUserId,
                                  ResponsibleDeptId,
                                  ResponsibleUserId,
                                  AqBarcodeStatus,
                                  LastModifiedBy,
                                  LastModifiedDate,
                                  AbnormalqualityId,
                                  AqBarcodeId
                              });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else //多筆異常回報
                        {
                            if (RepairCauseList.Length <= 0) throw new SystemException("對策資料不能為空,請重新確認");
                            List<string> QcItemList = new List<string>();

                            foreach (var item in RepairCauseList.Split(','))
                            {
                                QcItemList.Add(item);
                            }

                            foreach (var itemData in QcItemList)
                            {
                                string[] splitStr = { "%/%" }; //自行設定切割字串

                                #region //判斷品異條碼異常原因對策表是否存在
                                sql = @"SELECT AqQcItemId
                                        FROM QMS.AqQcItem
                                        WHERE AqQcItemId = @AqQcItemId";
                                dynamicParameters.Add("AqQcItemId", itemData.Split(splitStr, StringSplitOptions.RemoveEmptyEntries)[0]);

                                var resultAqQcItem = sqlConnection.Query(sql, dynamicParameters);
                                if (resultAqQcItem.Count() <= 0) throw new SystemException("品異條碼異常原因對策表資料不存在,請重新確認");
                                var AqQcItemId = Convert.ToInt32(itemData.Split(splitStr, StringSplitOptions.RemoveEmptyEntries)[0]);
                                #endregion

                                #region //判斷量測異常項目是否存在
                                sql = @"SELECT QcItemId
                                        FROM QMS.AqQcItem
                                        WHERE AqQcItemId = @AqQcItemId
                                        AND QcItemId = @QcItemId";
                                dynamicParameters.Add("AqQcItemId", itemData.Split(splitStr, StringSplitOptions.RemoveEmptyEntries)[0]);
                                dynamicParameters.Add("QcItemId", itemData.Split(splitStr, StringSplitOptions.RemoveEmptyEntries)[1]);

                                var resultQcItem = sqlConnection.Query(sql, dynamicParameters);
                                if (resultQcItem.Count() <= 0) throw new SystemException("量測異常項目資料不存在,請重新確認");
                                #endregion

                                #region //判斷不良原因代碼是否存在
                                sql = @"SELECT a.CauseId
                                        FROM QMS.DefectCause a 
                                        WHERE a.CauseId = @CauseId";
                                dynamicParameters.Add("CauseId", itemData.Split(splitStr, StringSplitOptions.RemoveEmptyEntries)[2]);

                                var resultDefectCause = sqlConnection.Query(sql, dynamicParameters);
                                if (resultDefectCause.Count() <= 0) throw new SystemException("不良原因代碼資料不存在,請重新確認");
                                foreach(var item in resultDefectCause) { }
                                var DefectCauseId = Convert.ToInt32(itemData.Split(splitStr, StringSplitOptions.RemoveEmptyEntries)[2]);
                                var DefectCauseDesc = itemData.Split(splitStr, StringSplitOptions.RemoveEmptyEntries)[3];
                                #endregion

                                #region //判斷對策代碼是否存在
                                sql = @"SELECT CauseId
                                    FROM QMS.RepairCause
                                    WHERE CauseNo = @CauseId";
                                dynamicParameters.Add("CauseId", itemData.Split(splitStr, StringSplitOptions.RemoveEmptyEntries)[4]);

                                var resultRepairCause = sqlConnection.Query(sql, dynamicParameters);
                                if (resultRepairCause.Count() <= 0) throw new SystemException("對策代碼資料不存在,請重新確認");
                                RepairCauseId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CauseId;
                                RepairCauseDesc = itemData.Split(splitStr, StringSplitOptions.RemoveEmptyEntries)[5];
                                #endregion

                                #region //更新-品異條碼異常原因對策表
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE QMS.AqQcItem SET
                                        DefectCauseId = @DefectCauseId,
                                        DefectCauseDesc = @DefectCauseDesc,
                                        RepairCauseId = @RepairCauseId,
                                        RepairCauseDesc = @RepairCauseDesc,
                                        LastModifiedBy = @LastModifiedBy,
                                        LastModifiedDate = @LastModifiedDate
                                        WHERE AqBarcodeId = @AqBarcodeId
                                        AND AqQcItemId = @AqQcItemId";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      DefectCauseId,
                                      DefectCauseDesc,
                                      RepairCauseId,
                                      RepairCauseDesc,
                                      LastModifiedBy,
                                      LastModifiedDate,
                                      AqBarcodeId,
                                      AqQcItemId
                                  });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                            }

                            #region //判斷對策代碼是否都有維護
                            int RepairCauseResult = -1;
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT a.RepairCauseId
                                    FROM QMS.AqQcItem a
                                    WHERE a.AqBarcodeId = @AqBarcodeId
                                    AND a.RepairCauseId is null";
                            dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                RepairCauseResult = 1;
                            }
                            else
                            {
                                RepairCauseResult = 2;
                            }
                            #endregion

                            #region //更新-異常條碼單據的填寫對策者
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.AqBarcode SET
                                    RepairCauseUserId = @RepairCauseUserId,
                                    ResponsibleDeptId = @ResponsibleDeptId,
                                    ResponsibleUserId = @ResponsibleUserId,
                                    AqBarcodeStatus = @AqBarcodeStatus,
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE AbnormalqualityId = @AbnormalqualityId
                                    AND AqBarcodeId = @AqBarcodeId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  RepairCauseUserId,
                                  ResponsibleDeptId,
                                  ResponsibleUserId,
                                  AqBarcodeStatus = "2",
                                  LastModifiedBy,
                                  LastModifiedDate,
                                  AbnormalqualityId,
                                  AqBarcodeId
                              });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
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

        #region //UpdateAqCountermeasureDetailBatch -- 更新品異單異常對策資料(批量) -- Shintokuro 2022.10.19
        public string UpdateAqCountermeasureDetailBatch(int AbnormalqualityId, string AqBarcodeIdList, int RepairCauseId, string RepairCauseDesc, int RepairCauseUserId)
        {
            try
            {
                if (AqBarcodeIdList.Length <= 0) throw new SystemException("【條碼列表】不能為空!");
                if (RepairCauseId <= 0) throw new SystemException("【對策代碼】不能為空!");
                if (RepairCauseUserId <= 0) throw new SystemException("【編寫對策者】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        string QcType = "";
                        foreach (var AqBarcodeId in AqBarcodeIdList.Split(','))
                        {
                            int? RepairCauseUserIdOriginal = -1;
                            int AqBarcodeStatusBase = -1;
                            #region //判斷異常條碼資料是否正確+判斷是否相同人修改
                            sql = @"SELECT a.RepairCauseUserId,a.AqBarcodeStatus,a.AqBarcodeId
                                    ,a1.QcType
                                    ,b.BarcodeNo
                                    FROM QMS.AqBarcode a
                                    INNER JOIN QMS.Abnormalquality a1 on a.AbnormalqualityId = a1.AbnormalqualityId
                                    LEFT JOIN MES.Barcode b on a.BarcodeId = b.BarcodeId
                                    WHERE a.AbnormalqualityId = @AbnormalqualityId
                                    AND a.AqBarcodeId = @AqBarcodeId";
                            dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                            dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("資料不存在,請重新確認");
                            foreach(var item in result)
                            {
                                RepairCauseUserIdOriginal = item.RepairCauseUserId;
                                AqBarcodeStatusBase = Convert.ToInt32(item.AqBarcodeStatus);
                                QcType = item.QcType;
                                if (QcType != "IQC")
                                {
                                    if (AqBarcodeStatusBase == 3) throw new SystemException("條碼:【" + item.BarcodeNo + "】該異常條碼已經完成對策確認,不可再更動");
                                    if (AqBarcodeStatusBase == 5) throw new SystemException("條碼:【" + item.BarcodeNo + "】該異常條碼已經完成會簽,不可再更動");
                                }
                                else
                                {
                                    if (AqBarcodeStatusBase == 3) throw new SystemException("該異常條碼已經完成對策確認,不可再更動");
                                    if (AqBarcodeStatusBase == 5) throw new SystemException("該異常條碼已經完成會簽,不可再更動");
                                }
                            }
                            //if (RepairCauseUserIdOriginal != null)
                            //{
                            //    if (RepairCauseUserIdOriginal != CreateBy) throw new SystemException("請不要修改他人資料!!!");
                            //}
                            #endregion

                            #region //判斷對策代碼是否存在
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                    FROM QMS.RepairCause a
                                    WHERE a.CauseId = @RepairCauseId";
                            dynamicParameters.Add("RepairCauseId", RepairCauseId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【對策代碼】不存在，請重新輸入!");
                            #endregion

                            #region //更新對策資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.AqBarcode SET
                                    RepairCauseId = @RepairCauseId,
                                    RepairCauseDesc = @RepairCauseDesc,
                                    RepairCauseUserId = @RepairCauseUserId,
                                    AqBarcodeStatus = @AqBarcodeStatus,
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE AbnormalqualityId = @AbnormalqualityId
                                    AND AqBarcodeId = @AqBarcodeId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  RepairCauseId,
                                  RepairCauseDesc,
                                  RepairCauseUserId,
                                  AqBarcodeStatus ="2",
                                  LastModifiedBy,
                                  LastModifiedDate,
                                  AbnormalqualityId,
                                  AqBarcodeId
                              });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
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

        #region //UpdateAqcConfirmDetail -- 更新品異單異常對策資料確認 -- Shintokuro 2022.10.12
        public string UpdateAqcConfirmDetail(int AbnormalqualityId, int AqBarcodeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        string JudgeConfirm = "";
                        int? RepairCauseUserId = 0;
                        int? ResponsibleSupervisorId = 0;
                        int AqBarcodeStatusBase = -1;

                        #region //判斷異常條碼資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 JudgeConfirm,RepairCauseUserId,ResponsibleSupervisorId,AqBarcodeStatus
                                FROM QMS.AqBarcode
                                WHERE AbnormalqualityId = @AbnormalqualityId
                                AND AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("資料不存在,請重新確認");
                        foreach(var item in result)
                        {
                            JudgeConfirm = item.JudgeConfirm;
                            RepairCauseUserId = item.RepairCauseUserId;
                            ResponsibleSupervisorId = item.ResponsibleSupervisorId;
                            AqBarcodeStatusBase = Convert.ToInt32(item.AqBarcodeStatus);

                        }
                        if (RepairCauseUserId == null) throw new SystemException("對策者未維護不能做對策確認,請重新確認");
                        if(ResponsibleSupervisorId != null) throw new SystemException("責任主管已經確認!!");
                        if (AqBarcodeStatusBase == 5) throw new SystemException("該異常條碼已經完成會簽,不可再更動");
                        #endregion


                        #region //更新責任主管資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.AqBarcode SET
                                ResponsibleSupervisorId = @ResponsibleSupervisorId,
                                LastModifiedBy = @LastModifiedBy,
                                AqBarcodeStatus = @AqBarcodeStatus,
                                LastModifiedDate = @LastModifiedDate
                                WHERE AbnormalqualityId = @AbnormalqualityId
                                AND AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              ResponsibleSupervisorId = CreateBy,
                              AqBarcodeStatus = '3',
                              LastModifiedBy,
                              LastModifiedDate,
                              AbnormalqualityId,
                              AqBarcodeId
                          });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "確認成功",
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

        #region //UpdateAqcConfirmDetailRe -- 更新品異單異常對策資料確認(撤銷) -- Shintokuro 2022.10.18
        public string UpdateAqcConfirmDetailRe(int AbnormalqualityId, int AqBarcodeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        int? nullData = null;
                        int AqBarcodeStatusBase = -1;
                        string AqBarcodeStatus = "";
                        #region //判斷異常條碼資料是否正確
                        int? JudgeUserId = -1;
                        sql = @"SELECT ResponsibleSupervisorId,JudgeUserId,AqBarcodeStatus
                                FROM QMS.AqBarcode
                                WHERE AbnormalqualityId = @AbnormalqualityId
                                AND AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("資料不存在,請重新確認");
                        foreach(var item in result)
                        {
                            JudgeUserId = item.JudgeUserId;
                            AqBarcodeStatusBase = Convert.ToInt32(item.AqBarcodeStatus);
                        }
                        if (AqBarcodeStatusBase == 5) throw new SystemException("該異常條碼已經完成會簽,不可再更動");

                        if (JudgeUserId > 0)
                        {
                            AqBarcodeStatus = "4";
                        }
                        else
                        {
                            AqBarcodeStatus = "2";
                        }
                        //int ResponsibleSupervisorIdOriginal = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ResponsibleSupervisorId;
                        //if (ResponsibleSupervisorIdOriginal != CreateBy) throw new SystemException("請不要修改他人資料!!!");
                        #endregion

                        #region //更新責任主管資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.AqBarcode SET
                                ResponsibleSupervisorId = @ResponsibleSupervisorId,
                                AqBarcodeStatus = @AqBarcodeStatus,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE AbnormalqualityId = @AbnormalqualityId
                                AND AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              ResponsibleSupervisorId = nullData,
                              AqBarcodeStatus ="2",
                              LastModifiedBy,
                              LastModifiedDate,
                              AbnormalqualityId,
                              AqBarcodeId
                          });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "撤銷成功",
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

        #region //UpdateAqcConfirmDetailBatch -- 更新品異單異常對策資料確認(批量) -- Shintokuro 2022.10.18
        public string UpdateAqcConfirmDetailBatch(int AbnormalqualityId, string AqBarcodeIdList)
        {
            try
            {
                if (AqBarcodeIdList.Length <= 0) throw new SystemException("【條碼列表】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        int? nullData = null;
                        string QcType = "";
                        foreach (var AqBarcodeId in AqBarcodeIdList.Split(',')){

                            int? RepairCauseUserId = -1;
                            int? ResponsibleSupervisorId = -1;
                            int AqBarcodeStatusBase = -1;

                            #region //判斷異常條碼資料是否正確
                            sql = @"SELECT a.RepairCauseUserId,a.ResponsibleSupervisorId,a.AqBarcodeStatus
                                    ,a1.QcType
                                    FROM QMS.AqBarcode a
                                    INNER JOIN QMS.Abnormalquality a1 on a.AbnormalqualityId = a1.AbnormalqualityId
                                    WHERE a.AbnormalqualityId = @AbnormalqualityId
                                    AND a.AqBarcodeId = @AqBarcodeId";
                            dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                            dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("資料不存在,請重新確認");
                            foreach(var item in result)
                            {
                                RepairCauseUserId = item.RepairCauseUserId;
                                ResponsibleSupervisorId = item.ResponsibleSupervisorId;
                                AqBarcodeStatusBase = Convert.ToInt32(item.AqBarcodeStatus);
                                QcType = item.QcType;
                                if (QcType != "IQC")
                                {
                                    if (RepairCauseUserId == null) throw new SystemException("條碼:【" + item.BarcodeNo + "】對策者未維護不能做對策確認,請重新確認");
                                    if (ResponsibleSupervisorId != null) throw new SystemException("條碼:【" + item.BarcodeNo + "】責任主管已經確認!!");
                                    if (AqBarcodeStatusBase == 5) throw new SystemException("條碼:【" + item.BarcodeNo + "】該異常條碼已經完成會簽,不可再更動");
                                }
                                else
                                {
                                    if (RepairCauseUserId == null) throw new SystemException("對策者未維護不能做對策確認,請重新確認");
                                    if (ResponsibleSupervisorId != null) throw new SystemException("責任主管已經確認!!");
                                    if (AqBarcodeStatusBase == 5) throw new SystemException("該異常條碼已經完成會簽,不可再更動");
                                }
                            }

                            #endregion

                            #region //更新責任主管資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.AqBarcode SET
                                    ResponsibleSupervisorId = @ResponsibleSupervisorId,
                                    AqBarcodeStatus = @AqBarcodeStatus,
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE AbnormalqualityId = @AbnormalqualityId
                                    AND AqBarcodeId = @AqBarcodeId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  ResponsibleSupervisorId = CreateBy,
                                  AqBarcodeStatus = '3',
                                  LastModifiedBy,
                                  LastModifiedDate,
                                  AbnormalqualityId,
                                  AqBarcodeId
                              });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "撤銷成功",
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

        #region //UpdateAqJudgmentDetail -- 更新品異單異常判定資料 -- Shintokuro 2022.10.13
        public string UpdateAqJudgmentDetail(int AbnormalqualityId, int AqBarcodeId, string JudgeStatus, int JudgeReturnMoProcessId
            , int JudgeReturnNextMoProcessId, string JudgeDesc, int JudgeUserId, string JudgeDate, int ResponsibleUserId, int ResponsibleDeptId
            , int ReleaseQty, string UserAqPhrase)
        {
            try
            {
                if (ResponsibleUserId <= 0) throw new SystemException("【責任者不能為空】,請重新確認");
                if (ResponsibleDeptId <= 0) throw new SystemException("【責任單位不能為空】,請重新確認");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        int? nullData = null;
                        bool lastMoProcess =false; //是否最後一站
                        bool OspMoProcess = false; //是否托外製程
                        int OspDetailId = -1; //托外單Id
                        int BarcodeProcessId = -1;
                        List<string> insertResult = new List<string>();
                        int? BarcodeId = -1;
                        string BarcodeNo = "";
                        string BarcodeStatus = "1";
                        int OspNextMoProcessId = -1;

                        if (JudgeStatus.Length <=0) throw new SystemException("請選擇判定結果");
                        if(JudgeStatus == "RW")
                        {
                            if(JudgeReturnMoProcessId <=0) throw new SystemException("判定結果為【返修】，需要選擇退回製程");
                        }
                        if (JudgeUserId <= 0) throw new SystemException("請選擇判定者");

                        #region //判斷是否相同人修改資料
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT JudgeUserId
                        //        FROM QMS.AqBarcode
                        //        WHERE AbnormalqualityId = @AbnormalqualityId
                        //        AND AqBarcodeId = @AqBarcodeId";
                        //dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                        //dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        //var result = sqlConnection.Query(sql, dynamicParameters);
                        //int? JudgeUserIdOriginal = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).JudgeUserId;

                        //if (JudgeUserIdOriginal != null)
                        //{
                        //    if (JudgeUserIdOriginal != CreateBy) throw new SystemException("請不要修改他人資料!!!");
                        //}
                        #endregion

                        #region //判斷責任單位是否存在
                        if(ResponsibleDeptId > 0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM BAS.Department a
                                WHERE a.DepartmentId = @ResponsibleDeptId";
                            dynamicParameters.Add("ResponsibleDeptId", ResponsibleDeptId);

                            var resultDepartment = sqlConnection.Query(sql, dynamicParameters);
                            if (resultDepartment.Count() <= 0) throw new SystemException("【責任單位】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //判斷責任者是否存在
                        if(ResponsibleUserId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ResponsibleUserId
                                AND a.Status = 'A'";
                            dynamicParameters.Add("ResponsibleUserId", ResponsibleUserId);

                            var resultUser = sqlConnection.Query(sql, dynamicParameters);
                            if (resultUser.Count() <= 0) throw new SystemException("【責任者】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //判斷異常條碼資料是否正確
                        int? MoId = -1;
                        int? GrDetailId = -1;
                        string QcType = "";
                        string JudgeConfirm = "";
                        int? MoProcessId = 0;
                        int? CurrentMoProcessId = 0;
                        int? RepairCauseUserId = 0;
                        int? ResponsibleSupervisorId = 0;
                        int AqBarcodeStatus = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.DefectCauseId,a.ResponsibleSupervisorId,a.AqBarcodeStatus,a.BarcodeId,a.JudgeConfirm,a.RepairCauseUserId,a.ResponsibleSupervisorId
                                ,b.MoProcessId,b.QcType,b.MoId,b.GrDetailId
                                ,c.BarcodeNo,c.CurrentMoProcessId
                                FROM QMS.AqBarcode a
                                INNER JOIN QMS.Abnormalquality b on a.AbnormalqualityId = b.AbnormalqualityId
                                LEFT JOIN MES.Barcode c on a.BarcodeId = c.BarcodeId
                                WHERE a.AbnormalqualityId = @AbnormalqualityId
                                AND a.AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("資料有問題,請重新輸入");
                        //int? ResponsibleSupervisorId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ResponsibleSupervisorId;
                        //if (ResponsibleSupervisorId == null) throw new SystemException("責任主管未確認不能判定!");
                        //if (AqBarcodeStatus != 3) throw new SystemException("異常條碼狀態為對策確認狀態才能做判定");
                        foreach(var item in result)
                        {
                            MoId = item.MoId;
                            GrDetailId = item.GrDetailId;
                            QcType = item.QcType;
                            BarcodeId = item.BarcodeId;
                            BarcodeNo = item.BarcodeNo;
                            MoProcessId = item.MoProcessId;
                            CurrentMoProcessId = item.CurrentMoProcessId;
                            JudgeConfirm = item.JudgeConfirm;
                            RepairCauseUserId = item.RepairCauseUserId;
                            ResponsibleSupervisorId = item.ResponsibleSupervisorId;
                            AqBarcodeStatus = Convert.ToInt32(item.AqBarcodeStatus);
                            if (AqBarcodeStatus == 5) throw new SystemException("該異常條碼已經完成會簽,不可再更動");
                            if(JudgeConfirm == "N")
                            {
                                if (item.DefectCauseId != null)
                                {
                                    if (item.RepairCauseUserId != null)
                                    {
                                        if (item.ResponsibleSupervisorId != null)
                                        {
                                            AqBarcodeStatus = 3;
                                        }
                                        else
                                        {
                                            AqBarcodeStatus = 2;
                                        }
                                    }
                                    else
                                    {
                                        AqBarcodeStatus = 4;
                                    }
                                }
                                else
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.DefectCauseId,a.RepairCauseId
                                            FROM QMS.AqQcItem a
                                            WHERE a.AqBarcodeId = @AqBarcodeId";
                                    dynamicParameters.Add("AqBarcodeId", AqBarcodeId);
                                    var resultAqQcItem = sqlConnection.Query(sql, dynamicParameters);
                                    foreach(var item1 in resultAqQcItem)
                                    {
                                        if (item1.DefectCauseId == null)
                                        {
                                            throw new SystemException("條碼:【" + item.BarcodeNo + "】必須處於回報完成狀態,才能做判定");

                                        }
                                    }
                                }
                            }
                        }

                        #endregion

                        if (QcType == "PVTQC")
                        {
                            if(JudgeStatus == "RW" && JudgeStatus == "SR") throw new SystemException("試產檢不能判定返修或是原站返修!!!");
                        }

                        #region //判斷該製程是否最後一站
                        if (QcType == "IPQC" && QcType == "NON")
                        {
                            int MaxSortNumber = -1;
                            int SortNumber = -1;
                            sql = @"SELECT c.MoId,b.MoProcessId,c.SortNumber,MAX(d1.SortNumber) MaxSortNumber
                                FROM QMS.AqBarcode a
                                INNER JOIN QMS.Abnormalquality b on a.AbnormalqualityId = b.AbnormalqualityId
                                INNER JOIN MES.MoProcess c on b.MoProcessId = c.MoProcessId
                                INNER JOIN MES.ManufactureOrder d on c.MoId = d.MoId
                                INNER JOIN MES.MoProcess d1 on d.MoId = d1.MoId
								WHERE a.AbnormalqualityId = @AbnormalqualityId
								GROUP BY c.MoId,b.MoProcessId,c.SortNumber";
                            dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("資料有問題,請重新輸入");
                            foreach(var item in result)
                            {
                                MaxSortNumber = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSortNumber;
                                SortNumber = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).SortNumber;
                            }
                            if (MaxSortNumber == SortNumber)
                            {
                                lastMoProcess = true;
                            }
                        }
                        #endregion

                        #region //判斷該製程是否委外製程
                        if (QcType != "PVTQC" && QcType != "IQC" && QcType != "CQAQC" && QcType != "MFGQC" && QcType != "TQC")
                        {
                            sql = @"SELECT TOP 1 a.Status
                                , b.BarcodeNo
                                , d.MoId, d.MoProcessId ,d.OspDetailId
                                FROM MES.OspReceiptBarcode a
                                INNER JOIN MES.OspBarcode b ON a.OspBarcodeId = b.OspBarcodeId
                                INNER JOIN MES.OspReceiptDetail c ON a.OsrDetailId = c.OsrDetailId
                                INNER JOIN MES.OspDetail d ON c.OspDetailId = d.OspDetailId
                                WHERE d.MoProcessId = @MoProcessId
                                AND b.BarcodeNo = @BarcodeNo
                                Order By c.CreateDate DESC";
                            dynamicParameters.Add("BarcodeNo", BarcodeNo);
                            if (QcType == "IPQC")
                            {
                                dynamicParameters.Add("MoProcessId", MoProcessId);
                            }
                            else if (QcType == "OQC")
                            {
                                dynamicParameters.Add("MoProcessId", CurrentMoProcessId);
                            }
                            else if(QcType == "NON")
                            {
                                dynamicParameters.Add("MoProcessId", MoProcessId);
                            }

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                OspMoProcess = true;
                                OspDetailId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).OspDetailId;
                            }
                        }
                        #endregion

                        #region //判斷退回站別是否存在
                        if (JudgeReturnMoProcessId > 0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                FROM MES.MoProcess a 
                                WHERE a.MoId = @MoId
                                AND a.MoProcessId = @JudgeReturnMoProcessId";
                            dynamicParameters.Add("MoId", MoId);
                            dynamicParameters.Add("JudgeReturnMoProcessId", JudgeReturnMoProcessId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("退回站別不存在，請重新輸入!");
                        }
                        #endregion

                        #region //撈取BarcodeProcessId
                        if (QcType != "PVTQC" && QcType != "IQC" && QcType != "CQAQC" && QcType != "MFGQC" && QcType != "TQC")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.BarcodeProcessId
                                FROM MES.BarcodeProcess a
                                WHERE a.MoProcessId = @MoProcessId
                                AND a.BarcodeId = @BarcodeId
                                AND a.StartDate = (SELECT MAX(z.StartDate)
                                                                FROM MES.BarcodeProcess z
                                                               WHERE z.BarcodeId = @BarcodeId
                                                                 AND z.MoProcessId = @MoProcessId)";
                            dynamicParameters.Add("BarcodeId", BarcodeId);
                            if (QcType == "IPQC")
                            {
                                dynamicParameters.Add("MoProcessId", MoProcessId);
                            }
                            else if (QcType == "OQC")
                            {
                                dynamicParameters.Add("MoProcessId", CurrentMoProcessId);
                            }
                            else if (QcType == "NON")
                            {
                                dynamicParameters.Add("MoProcessId", MoProcessId);
                            }
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("資料有問題,請重新輸入");
                            BarcodeProcessId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).BarcodeProcessId;
                            insertResult.Add(JudgeStatus);
                            insertResult.Add(Convert.ToString(BarcodeProcessId));

                        }
                        #endregion

                        #region //更新判定資料

                        #region //更新Barcode NextMoProcessId
                        if (JudgeConfirm == "N")
                        {
                            //判定狀態為報廢時, 條碼流程為0, NextMoProcessId = -1

                            switch (QcType)
                            {
                                case "IQC":
                                    #region //撈取進貨數量
                                    int ReceiptQty = 0;
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ReceiptQty
                                        FROM SCM.GrDetail a
                                        WHERE a.GrDetailId = @GrDetailId";
                                    dynamicParameters.Add("GrDetailId", GrDetailId);

                                    var resultReceiptQty = sqlConnection.Query(sql, dynamicParameters);
                                    foreach (var item in resultReceiptQty)
                                    {
                                        ReceiptQty = item.ReceiptQty;
                                    }
                                    if (ReceiptQty < ReleaseQty) throw new SystemException("特採數量不可超過進貨單進貨數" + ReceiptQty + "!!!");
                                    if (ReleaseQty < 0) throw new SystemException("特採數量不可為負!!!");
                                    #endregion

                                    switch (JudgeStatus)
                                    {
                                        case "R":
                                            #region //更新進貨單單身資料
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE SCM.GrDetail SET
                                                QcStatus = @QcStatus,
                                                AcceptQty = @ReleaseQty,
                                                AvailableQty = @ReleaseQty,
                                                ReturnQty = ReceiptQty - @ReleaseQty,
                                                LastModifiedBy = @LastModifiedBy,
                                                LastModifiedDate = @LastModifiedDate
                                                WHERE GrDetailId = @GrDetailId";
                                            dynamicParameters.AddDynamicParams(
                                              new
                                              {
                                                  QcStatus = "4",
                                                  ReleaseQty = ReleaseQty,
                                                  LastModifiedBy,
                                                  LastModifiedDate,
                                                  GrDetailId
                                              });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion
                                            break;
                                        case "S":
                                            #region //更新進貨單單身資料
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE SCM.GrDetail SET
                                                QcStatus = @QcStatus,
                                                AcceptQty = @AcceptQty,
                                                AvailableQty = @AvailableQty,
                                                ReturnQty = ReceiptQty,
                                                LastModifiedBy = @LastModifiedBy,
                                                LastModifiedDate = @LastModifiedDate
                                                WHERE GrDetailId = @GrDetailId";
                                            dynamicParameters.AddDynamicParams(
                                              new
                                              {
                                                  QcStatus = "3",
                                                  AcceptQty = 0,
                                                  AvailableQty = 0,
                                                  LastModifiedBy,
                                                  LastModifiedDate,
                                                  GrDetailId
                                              });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion

                                            break;
                                        case "AM":
                                            break;
                                    }
                                    break;
                                case "CQAQC":
                                case "MFGQC":
                                case "TQC":
                                    break;
                                default:
                                    switch (JudgeStatus)
                                    {
                                        #region //報廢
                                        case "S":
                                            #region 製程為托外時,修改MES.OspReceiptBarcode資料 托外條碼狀態
                                            if (OspMoProcess == true)
                                            {
                                                //修改 托外條碼狀態
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE a 
                                            SET
                                            a.Status = @JudgeStatus,
                                            a.LastModifiedBy = @LastModifiedBy,
                                            a.LastModifiedDate = @LastModifiedDate
                                            FROM MES.OspReceiptBarcode a
                                            INNER JOIN MES.OspBarcode b on a.OspBarcodeId = b.OspBarcodeId
                                            WHERE b.BarcodeNo = @BarcodeNo";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      JudgeStatus,
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      BarcodeNo
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                                //修改 MES.OspReceiptDetail資料 托外入庫詳細資料數量
                                                dynamicParameters = new DynamicParameters();

                                                sql = @"UPDATE MES.OspReceiptDetail SET
                                                    AvailableQty = AvailableQty - 1,
                                                    ReturnQty = ReturnQty + 1,
                                                    LastModifiedBy = @LastModifiedBy,
                                                    LastModifiedDate = @LastModifiedDate
                                                    WHERE OspDetailId = @OspDetailId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      OspDetailId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            }
                                            #endregion

                                            if (QcType != "PVTQC")
                                            {
                                                #region //修改MES.Barcode的資料 狀態+下一站
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE MES.Barcode SET
                                        BarcodeStatus = '0',
                                        NextMoProcessId = -1,
                                        CurrentProdStatus = @JudgeStatus,
                                        LastModifiedBy = @LastModifiedBy,
                                        LastModifiedDate = @LastModifiedDate
                                        WHERE BarcodeId = @BarcodeId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      JudgeStatus,
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      BarcodeId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                                #endregion

                                                #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                                                UpdateBarcodeProcessStatus("U", BarcodeProcessId, "S");
                                                #endregion
                                            }

                                            break;
                                        #endregion
                                        #region //返修
                                        case "RW":
                                            #region 製程為托外時,修改MES.OspReceiptBarcode資料 托外條碼狀態
                                            if (OspMoProcess == true)
                                            {
                                                //修改 托外條碼狀態
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE a 
                                            SET
                                            a.Status = @JudgeStatus,
                                            a.LastModifiedBy = @LastModifiedBy,
                                            a.LastModifiedDate = @LastModifiedDate
                                            FROM MES.OspReceiptBarcode a
                                            INNER JOIN MES.OspBarcode b on a.OspBarcodeId = b.OspBarcodeId
                                            WHERE b.OspBarcodeId = @BarcodeId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      JudgeStatus,
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      BarcodeId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                                //修改 托外入庫詳細資料數量
                                                dynamicParameters = new DynamicParameters();

                                                sql = @"UPDATE MES.OspReceiptDetail SET
                                                    AvailableQty = AvailableQty - 1,
                                                    ReturnQty = ReturnQty + 1,
                                                    LastModifiedBy = @LastModifiedBy,
                                                    LastModifiedDate = @LastModifiedDate
                                                    WHERE OspDetailId = @OspDetailId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      OspDetailId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            }
                                            #endregion

                                            #region //修改MES.Barcode的資料 狀態+下一站
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE MES.Barcode SET
                                        NextMoProcessId = @JudgeReturnMoProcessId,
                                        CurrentProdStatus = @JudgeStatus,
                                        BarcodeStatus='1',
                                        LastModifiedBy = @LastModifiedBy,
                                        LastModifiedDate = @LastModifiedDate
                                        WHERE BarcodeId = @BarcodeId";
                                            dynamicParameters.AddDynamicParams(
                                              new
                                              {
                                                  JudgeReturnMoProcessId,
                                                  JudgeStatus,
                                                  LastModifiedBy,
                                                  LastModifiedDate,
                                                  BarcodeId
                                              });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion

                                            #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                                            UpdateBarcodeProcessStatus("U", BarcodeProcessId, "F");
                                            #endregion
                                            break;
                                        #endregion
                                        #region //當站返修
                                        case "SR":
                                            #region 製程為托外時,修改MES.OspReceiptBarcode資料 托外條碼狀態
                                            if (OspMoProcess == true)
                                            {
                                                //修改 托外條碼狀態
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE a 
                                            SET
                                            a.Status = @JudgeStatus,
                                            a.LastModifiedBy = @LastModifiedBy,
                                            a.LastModifiedDate = @LastModifiedDate
                                            FROM MES.OspReceiptBarcode a
                                            INNER JOIN MES.OspBarcode b on a.OspBarcodeId = b.OspBarcodeId
                                            WHERE b.OspBarcodeId = @BarcodeId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      JudgeStatus,
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      BarcodeId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                                //修改 MES.OspReceiptDetail資料 托外入庫詳細資料數量
                                                dynamicParameters = new DynamicParameters();

                                                sql = @"UPDATE MES.OspReceiptDetail SET
                                                    AvailableQty = AvailableQty - 1,
                                                    ReturnQty = ReturnQty + 1,
                                                    LastModifiedBy = @LastModifiedBy,
                                                    LastModifiedDate = @LastModifiedDate
                                                    WHERE OspDetailId = @OspDetailId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      OspDetailId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            }
                                            #endregion

                                            #region //修改MES.Barcode的資料 狀態+下一站
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE MES.Barcode SET
                                            BarcodeStatus='1',
                                            NextMoProcessId = CurrentMoProcessId,
                                            CurrentProdStatus = @JudgeStatus,
                                            LastModifiedBy = @LastModifiedBy,
                                            LastModifiedDate = @LastModifiedDate
                                            WHERE BarcodeId = @BarcodeId";
                                            dynamicParameters.AddDynamicParams(
                                              new
                                              {
                                                  JudgeStatus,
                                                  LastModifiedBy,
                                                  LastModifiedDate,
                                                  BarcodeId
                                              });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion

                                            #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                                            UpdateBarcodeProcessStatus("U", BarcodeProcessId, "F");
                                            #endregion
                                            break;
                                        #endregion
                                        #region //誤判 / 特採 / 不良品/流程更換
                                        case "R": //特採放行
                                        case "M": //誤判
                                        case "NK": //不良品/流程更換
                                            #region //判斷是否最後一站
                                            if (QcType == "IPQC" || QcType == "NON")
                                            {
                                                if (lastMoProcess == true)
                                                {
                                                    BarcodeStatus = "0";
                                                }
                                            }
                                            #endregion

                                            #region //不良品/流程更換=>採用強制結束狀態
                                            if (JudgeStatus == "NK")
                                            {
                                                BarcodeStatus = "7";
                                            }
                                            #endregion


                                            if (QcType != "PVTQC")
                                            {
                                                #region //修改條碼狀態
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE MES.Barcode SET
                                            CurrentProdStatus = @JudgeStatus,
                                            BarcodeStatus = @BarcodeStatus,
                                            LastModifiedBy = @LastModifiedBy,
                                            LastModifiedDate = @LastModifiedDate
                                            WHERE BarcodeId = @BarcodeId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      JudgeStatus,
                                                      BarcodeStatus,
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      BarcodeId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                                #endregion

                                                #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                                                UpdateBarcodeProcessStatus("U", BarcodeProcessId, "P");
                                                #endregion

                                            }
                                            break;
                                        #endregion

                                        default: throw new SystemException("判定結果未設定!");

                                    }
                                    break;
                            }

                            //if(QcType != "IQC")
                            //{
                            //    switch (JudgeStatus)
                            //    {
                            //        #region //報廢
                            //        case "S":
                            //            #region 製程為托外時,修改MES.OspReceiptBarcode資料 托外條碼狀態
                            //            if (OspMoProcess == true)
                            //            {
                            //                //修改 托外條碼狀態
                            //                dynamicParameters = new DynamicParameters();
                            //                sql = @"UPDATE a 
                            //                SET
                            //                a.Status = @JudgeStatus,
                            //                a.LastModifiedBy = @LastModifiedBy,
                            //                a.LastModifiedDate = @LastModifiedDate
                            //                FROM MES.OspReceiptBarcode a
                            //                INNER JOIN MES.OspBarcode b on a.OspBarcodeId = b.OspBarcodeId
                            //                WHERE b.BarcodeNo = @BarcodeNo";
                            //                dynamicParameters.AddDynamicParams(
                            //                  new
                            //                  {
                            //                      JudgeStatus,
                            //                      LastModifiedBy,
                            //                      LastModifiedDate,
                            //                      BarcodeNo
                            //                  });
                            //                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            //                //修改 MES.OspReceiptDetail資料 托外入庫詳細資料數量
                            //                dynamicParameters = new DynamicParameters();

                            //                sql = @"UPDATE MES.OspReceiptDetail SET
                            //                        AvailableQty = AvailableQty - 1,
                            //                        ReturnQty = ReturnQty + 1,
                            //                        LastModifiedBy = @LastModifiedBy,
                            //                        LastModifiedDate = @LastModifiedDate
                            //                        WHERE OspDetailId = @OspDetailId";
                            //                dynamicParameters.AddDynamicParams(
                            //                  new
                            //                  {
                            //                      LastModifiedBy,
                            //                      LastModifiedDate,
                            //                      OspDetailId
                            //                  });
                            //                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //            }
                            //            #endregion

                            //            if (QcType != "PVTQC")
                            //            {
                            //                #region //修改MES.Barcode的資料 狀態+下一站
                            //                dynamicParameters = new DynamicParameters();
                            //                sql = @"UPDATE MES.Barcode SET
                            //            BarcodeStatus = '0',
                            //            NextMoProcessId = -1,
                            //            CurrentProdStatus = @JudgeStatus,
                            //            LastModifiedBy = @LastModifiedBy,
                            //            LastModifiedDate = @LastModifiedDate
                            //            WHERE BarcodeId = @BarcodeId";
                            //                dynamicParameters.AddDynamicParams(
                            //                  new
                            //                  {
                            //                      JudgeStatus,
                            //                      LastModifiedBy,
                            //                      LastModifiedDate,
                            //                      BarcodeId
                            //                  });
                            //                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //                #endregion

                            //                #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                            //                UpdateBarcodeProcessStatus("U", BarcodeProcessId, "S");
                            //                #endregion
                            //            }

                            //            break;
                            //        #endregion
                            //        #region //返修
                            //        case "RW":
                            //            #region 製程為托外時,修改MES.OspReceiptBarcode資料 托外條碼狀態
                            //            if (OspMoProcess == true)
                            //            {
                            //                //修改 托外條碼狀態
                            //                dynamicParameters = new DynamicParameters();
                            //                sql = @"UPDATE a 
                            //                SET
                            //                a.Status = @JudgeStatus,
                            //                a.LastModifiedBy = @LastModifiedBy,
                            //                a.LastModifiedDate = @LastModifiedDate
                            //                FROM MES.OspReceiptBarcode a
                            //                INNER JOIN MES.OspBarcode b on a.OspBarcodeId = b.OspBarcodeId
                            //                WHERE b.OspBarcodeId = @BarcodeId";
                            //                dynamicParameters.AddDynamicParams(
                            //                  new
                            //                  {
                            //                      JudgeStatus,
                            //                      LastModifiedBy,
                            //                      LastModifiedDate,
                            //                      BarcodeId
                            //                  });
                            //                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            //                //修改 托外入庫詳細資料數量
                            //                dynamicParameters = new DynamicParameters();

                            //                sql = @"UPDATE MES.OspReceiptDetail SET
                            //                        AvailableQty = AvailableQty - 1,
                            //                        ReturnQty = ReturnQty + 1,
                            //                        LastModifiedBy = @LastModifiedBy,
                            //                        LastModifiedDate = @LastModifiedDate
                            //                        WHERE OspDetailId = @OspDetailId";
                            //                dynamicParameters.AddDynamicParams(
                            //                  new
                            //                  {
                            //                      LastModifiedBy,
                            //                      LastModifiedDate,
                            //                      OspDetailId
                            //                  });
                            //                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //            }
                            //            #endregion

                            //            #region //修改MES.Barcode的資料 狀態+下一站
                            //            dynamicParameters = new DynamicParameters();
                            //            sql = @"UPDATE MES.Barcode SET
                            //            NextMoProcessId = @JudgeReturnMoProcessId,
                            //            CurrentProdStatus = @JudgeStatus,
                            //            BarcodeStatus='1',
                            //            LastModifiedBy = @LastModifiedBy,
                            //            LastModifiedDate = @LastModifiedDate
                            //            WHERE BarcodeId = @BarcodeId";
                            //            dynamicParameters.AddDynamicParams(
                            //              new
                            //              {
                            //                  JudgeReturnMoProcessId,
                            //                  JudgeStatus,
                            //                  LastModifiedBy,
                            //                  LastModifiedDate,
                            //                  BarcodeId
                            //              });
                            //            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //            #endregion

                            //            #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                            //            UpdateBarcodeProcessStatus("U", BarcodeProcessId, "F");
                            //            #endregion
                            //            break;
                            //        #endregion
                            //        #region //當站返修
                            //        case "SR":
                            //            #region 製程為托外時,修改MES.OspReceiptBarcode資料 托外條碼狀態
                            //            if (OspMoProcess == true)
                            //            {
                            //                //修改 托外條碼狀態
                            //                dynamicParameters = new DynamicParameters();
                            //                sql = @"UPDATE a 
                            //                SET
                            //                a.Status = @JudgeStatus,
                            //                a.LastModifiedBy = @LastModifiedBy,
                            //                a.LastModifiedDate = @LastModifiedDate
                            //                FROM MES.OspReceiptBarcode a
                            //                INNER JOIN MES.OspBarcode b on a.OspBarcodeId = b.OspBarcodeId
                            //                WHERE b.OspBarcodeId = @BarcodeId";
                            //                dynamicParameters.AddDynamicParams(
                            //                  new
                            //                  {
                            //                      JudgeStatus,
                            //                      LastModifiedBy,
                            //                      LastModifiedDate,
                            //                      BarcodeId
                            //                  });
                            //                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            //                //修改 MES.OspReceiptDetail資料 托外入庫詳細資料數量
                            //                dynamicParameters = new DynamicParameters();

                            //                sql = @"UPDATE MES.OspReceiptDetail SET
                            //                        AvailableQty = AvailableQty - 1,
                            //                        ReturnQty = ReturnQty + 1,
                            //                        LastModifiedBy = @LastModifiedBy,
                            //                        LastModifiedDate = @LastModifiedDate
                            //                        WHERE OspDetailId = @OspDetailId";
                            //                dynamicParameters.AddDynamicParams(
                            //                  new
                            //                  {
                            //                      LastModifiedBy,
                            //                      LastModifiedDate,
                            //                      OspDetailId
                            //                  });
                            //                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //            }
                            //            #endregion

                            //            #region //修改MES.Barcode的資料 狀態+下一站
                            //            dynamicParameters = new DynamicParameters();
                            //            sql = @"UPDATE MES.Barcode SET
                            //                BarcodeStatus='1',
                            //                NextMoProcessId = CurrentMoProcessId,
                            //                CurrentProdStatus = @JudgeStatus,
                            //                LastModifiedBy = @LastModifiedBy,
                            //                LastModifiedDate = @LastModifiedDate
                            //                WHERE BarcodeId = @BarcodeId";
                            //            dynamicParameters.AddDynamicParams(
                            //              new
                            //              {
                            //                  JudgeStatus,
                            //                  LastModifiedBy,
                            //                  LastModifiedDate,
                            //                  BarcodeId
                            //              });
                            //            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //            #endregion

                            //            #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                            //            UpdateBarcodeProcessStatus("U", BarcodeProcessId, "F");
                            //            #endregion
                            //            break;
                            //        #endregion
                            //        #region //誤判 / 特採 / 不良品/流程更換
                            //        case "R": //特採放行
                            //        case "M": //誤判
                            //        case "NK": //不良品/流程更換
                            //            #region //判斷是否最後一站
                            //            if (QcType == "IPQC" || QcType == "NON")
                            //            {
                            //                if (lastMoProcess == true)
                            //                {
                            //                    BarcodeStatus = "0";
                            //                }
                            //            }
                            //            #endregion

                            //            #region //不良品/流程更換=>採用強制結束狀態
                            //            if (JudgeStatus == "NK")
                            //            {
                            //                BarcodeStatus = "7";
                            //            }
                            //            #endregion


                            //            if (QcType != "PVTQC")
                            //            {
                            //                #region //修改條碼狀態
                            //                dynamicParameters = new DynamicParameters();
                            //                sql = @"UPDATE MES.Barcode SET
                            //                CurrentProdStatus = @JudgeStatus,
                            //                BarcodeStatus = @BarcodeStatus,
                            //                LastModifiedBy = @LastModifiedBy,
                            //                LastModifiedDate = @LastModifiedDate
                            //                WHERE BarcodeId = @BarcodeId";
                            //                dynamicParameters.AddDynamicParams(
                            //                  new
                            //                  {
                            //                      JudgeStatus,
                            //                      BarcodeStatus,
                            //                      LastModifiedBy,
                            //                      LastModifiedDate,
                            //                      BarcodeId
                            //                  });
                            //                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //                #endregion

                            //                #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                            //                UpdateBarcodeProcessStatus("U", BarcodeProcessId, "P");
                            //                #endregion

                            //            }
                            //            break;
                            //        #endregion

                            //        default: throw new SystemException("判定結果未設定!");

                            //    }
                            //}
                            //else
                            //{
                            //    #region //撈取進貨數量
                            //    int ReceiptQty = 0;
                            //    dynamicParameters = new DynamicParameters();
                            //    sql = @"SELECT ReceiptQty
                            //            FROM SCM.GrDetail a
                            //            WHERE a.GrDetailId = @GrDetailId";
                            //    dynamicParameters.Add("GrDetailId", GrDetailId);

                            //    var resultReceiptQty = sqlConnection.Query(sql, dynamicParameters);
                            //    foreach(var item in resultReceiptQty)
                            //    {
                            //        ReceiptQty = item.ReceiptQty;
                            //    }
                            //    if(ReceiptQty< ReleaseQty) throw new SystemException("特採數量不可超過進貨單進貨數"+ ReceiptQty + "!!!");
                            //    if(ReleaseQty < 0) throw new SystemException("特採數量不可為負!!!");
                            //    #endregion

                            //    switch (JudgeStatus)
                            //    {
                            //        case "R":
                            //            #region //更新進貨單單身資料
                            //            dynamicParameters = new DynamicParameters();
                            //            sql = @"UPDATE SCM.GrDetail SET
                            //                    QcStatus = @QcStatus,
                            //                    AcceptQty = @ReleaseQty,
                            //                    AvailableQty = @ReleaseQty,
                            //                    ReturnQty = ReceiptQty - @ReleaseQty,
                            //                    LastModifiedBy = @LastModifiedBy,
                            //                    LastModifiedDate = @LastModifiedDate
                            //                    WHERE GrDetailId = @GrDetailId";
                            //            dynamicParameters.AddDynamicParams(
                            //              new
                            //              {
                            //                  QcStatus = "4",
                            //                  ReleaseQty = ReleaseQty,
                            //                  LastModifiedBy,
                            //                  LastModifiedDate,
                            //                  GrDetailId
                            //              });
                            //            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //            #endregion
                            //            break;
                            //        case "S":
                            //            #region //更新進貨單單身資料
                            //            dynamicParameters = new DynamicParameters();
                            //            sql = @"UPDATE SCM.GrDetail SET
                            //                    QcStatus = @QcStatus,
                            //                    AcceptQty = @AcceptQty,
                            //                    AvailableQty = @AvailableQty,
                            //                    ReturnQty = ReceiptQty,
                            //                    LastModifiedBy = @LastModifiedBy,
                            //                    LastModifiedDate = @LastModifiedDate
                            //                    WHERE GrDetailId = @GrDetailId";
                            //            dynamicParameters.AddDynamicParams(
                            //              new
                            //              {
                            //                  QcStatus = "3",
                            //                  AcceptQty = 0,
                            //                  AvailableQty = 0,
                            //                  LastModifiedBy,   
                            //                  LastModifiedDate,
                            //                  GrDetailId
                            //              });
                            //            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //            #endregion

                            //            break;
                            //        case "AM":
                            //            break;
                            //    }
                            //}

                            #region //更新品異單資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.AqBarcode SET
                                    JudgeStatus = @JudgeStatus,
                                    JudgeReturnMoProcessId = @JudgeReturnMoProcessId,
                                    JudgeReturnNextMoProcessId = @JudgeReturnNextMoProcessId,
                                    JudgeDesc = @JudgeDesc,
                                    JudgeUserId = @JudgeUserId,
                                    JudgeConfirm = @JudgeConfirm,
                                    JudgeDate = @JudgeDate,
                                    ResponsibleUserId = @ResponsibleUserId,
                                    ResponsibleDeptId = @ResponsibleDeptId,
                                    AqBarcodeStatus = @AqBarcodeStatus,
                                    ReleaseQty = @ReleaseQty,
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE AbnormalqualityId = @AbnormalqualityId
                                    AND AqBarcodeId = @AqBarcodeId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  JudgeStatus,
                                  JudgeReturnMoProcessId = JudgeReturnMoProcessId > 0 ? JudgeReturnMoProcessId : nullData,
                                  JudgeReturnNextMoProcessId = JudgeReturnNextMoProcessId > 0 ? JudgeReturnNextMoProcessId : nullData,
                                  JudgeDesc,
                                  JudgeUserId,
                                  JudgeConfirm = 'Y',
                                  AqBarcodeStatus,
                                  JudgeDate,
                                  ResponsibleUserId,
                                  ResponsibleDeptId,
                                  ReleaseQty = QcType == "IQC"? ReleaseQty : (int?)null,
                                  LastModifiedBy,
                                  LastModifiedDate,
                                  AbnormalqualityId,
                                  AqBarcodeId
                              });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            #region //更新品異單資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.AqBarcode SET
                                    ResponsibleUserId = @ResponsibleUserId,
                                    ResponsibleDeptId = @ResponsibleDeptId,
                                    JudgeDesc = @JudgeDesc,
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE AbnormalqualityId = @AbnormalqualityId
                                    AND AqBarcodeId = @AqBarcodeId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  ResponsibleUserId,
                                  ResponsibleDeptId,
                                  JudgeDesc,
                                  LastModifiedBy,
                                  LastModifiedDate,
                                  AbnormalqualityId,
                                  AqBarcodeId
                              });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #endregion

                        #region //判定原因是否納入常用片語
                        if (UserAqPhrase == "Y")
                        {
                            #region //判段該原因是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM QMS.AqPhrase a
                                    WHERE a.PhraseText = @PhraseText";
                            dynamicParameters.Add("PhraseText", JudgeDesc);

                            var resultDepartment = sqlConnection.Query(sql, dynamicParameters);
                            if (resultDepartment.Count() <= 0)
                            {
                                #region 常用片語建立
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO QMS.AqPhrase (CompanyId, PhraseText,
                                    CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @PhraseText,
                                    @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        CompanyId = CurrentCompany,
                                        PhraseText = JudgeDesc,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion
                        }
                        #endregion

                        if (QcType == "IQC")
                        {
                            #region //寄送品異單簽核完成信
                            SendIQcAqConfirmMail(sqlConnection, AbnormalqualityId);

                            #endregion
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult

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

        #region //UpdateAqJudgmentDetailBatch -- 更新品異單異常判定資料(批量) -- Shintokuro 2022.10.19
        public string UpdateAqJudgmentDetailBatch(int AbnormalqualityId, string AqBarcodeIdList, string JudgeStatus, int JudgeReturnMoProcessId
            , int JudgeReturnNextMoProcessId, string JudgeDesc, int JudgeUserId, string JudgeDate, int ResponsibleDeptId, int ResponsibleUserId
            , int ReleaseQty, string UserAqPhrase)
        {
            try
            {
                if (AqBarcodeIdList.Length <= 0) throw new SystemException("【條碼列表】不能為空!");
                if (JudgeStatus.Length <= 0) throw new SystemException("請選擇判定結果");
                if (JudgeUserId <= 0) throw new SystemException("請選擇判定者");
                if (JudgeStatus == "RW")
                {
                    if (JudgeReturnMoProcessId <= 0) throw new SystemException("判定結果為【返修】，需要選擇退回製程");
                }
                if (ResponsibleDeptId > 0)
                {
                    if(ResponsibleUserId <= 0) throw new SystemException("如果有填寫責任單位必續填寫責任者,請重新確認");
                }
                if (ResponsibleUserId > 0)
                {
                    if (ResponsibleDeptId <= 0) throw new SystemException("如果有填寫責任者必續填寫責任單位,請重新確認");
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        int? nullData = null;
                        bool lastMoProcess = false; //是否最後一站
                        bool OspMoProcess = false; //是否托外製程
                        int OspDetailId = -1; //托外單Id
                        int BarcodeProcessId = -1;
                        List<string> insertResult = new List<string>();
                        int? BarcodeId = -1;
                        string BarcodeNo = "";
                        string BarcodeStatus = "0";
                        string JudgeConfirm = "";
                        int OspNextMoProcessId = -1;

                        #region //判斷責任單位是否存在
                        if (ResponsibleDeptId > 0)
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT TOP 1 1
                                    FROM BAS.Department a
                                    WHERE a.DepartmentId = @ResponsibleDeptId";
                            dynamicParameters.Add("ResponsibleDeptId", ResponsibleDeptId);

                            var resultDepartment = sqlConnection.Query(sql, dynamicParameters);
                            if (resultDepartment.Count() <= 0) throw new SystemException("【責任單位】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //判斷責任者是否存在
                        if (ResponsibleUserId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User] a
                                    WHERE a.UserId = @ResponsibleUserId
                                    AND a.Status = 'A'";
                            dynamicParameters.Add("ResponsibleUserId", ResponsibleUserId);

                            var resultUser = sqlConnection.Query(sql, dynamicParameters);
                            if (resultUser.Count() <= 0) throw new SystemException("【責任者】不存在，請重新輸入!");
                        }
                        #endregion

                        foreach (var AqBarcodeId in AqBarcodeIdList.Split(','))
                        {
                            #region //判斷異常條碼資料是否正確
                            int? ResponsibleSupervisorId = -1;
                            int? MoId = -1;
                            string QcType = "";
                            int? MoProcessId = -1;
                            int? CurrentMoProcessId = -1;
                            int AqBarcodeStatus = -1;
                            int? GrDetailId = -1;
                            sql = @"SELECT a.DefectCauseId,a.RepairCauseUserId,a.ResponsibleSupervisorId,a.AqBarcodeStatus,a.BarcodeId,a.JudgeConfirm
                                    ,b.MoProcessId,b.QcType,b.MoId,b.GrDetailId
                                    ,c.BarcodeNo,c.CurrentMoProcessId
                                    FROM QMS.AqBarcode a
                                    INNER JOIN QMS.Abnormalquality b on a.AbnormalqualityId = b.AbnormalqualityId
                                    LEFT JOIN MES.Barcode c on a.BarcodeId = c.BarcodeId
                                    WHERE a.AbnormalqualityId = @AbnormalqualityId
                                    AND a.AqBarcodeId = @AqBarcodeId";
                            dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                            dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【品異單】異常條碼資料有問題,請重新輸入");
                            foreach(var item in result)
                            {
                                ResponsibleSupervisorId = item.ResponsibleSupervisorId;
                                MoId = item.MoId;
                                GrDetailId = item.GrDetailId;
                                QcType = item.QcType;
                                BarcodeId = item.BarcodeId;
                                BarcodeNo = item.BarcodeNo;
                                if(item.QcType == "OQC" || item.QcType == "IQC")
                                {
                                    MoProcessId = null;
                                }
                                else
                                {
                                    MoProcessId =item.MoProcessId != null ? item.MoProcessId : throw new SystemException("條碼的異常製成不存在,請重新確認"); 
                                }
                                CurrentMoProcessId = item.CurrentMoProcessId;
                                AqBarcodeStatus = Convert.ToInt32(item.AqBarcodeStatus);
                                if (AqBarcodeStatus == 5) throw new SystemException("該異常條碼已經完成會簽,不可再更動");

                                JudgeConfirm = item.JudgeConfirm;
                                if (JudgeConfirm == "N")
                                {
                                    if (item.DefectCauseId != null)
                                    {
                                        if (item.RepairCauseUserId != null)
                                        {
                                            if (item.ResponsibleSupervisorId != null)
                                            {
                                                AqBarcodeStatus = 3;
                                            }
                                            else
                                            {
                                                AqBarcodeStatus = 2;
                                            }
                                        }
                                        else
                                        {
                                            AqBarcodeStatus = 4;
                                        }
                                    }
                                    else
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT a.DefectCauseId,a.RepairCauseId
                                            FROM QMS.AqQcItem a
                                            WHERE a.AqBarcodeId = @AqBarcodeId";
                                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);
                                        var resultAqQcItem = sqlConnection.Query(sql, dynamicParameters);
                                        foreach (var item1 in resultAqQcItem)
                                        {
                                            if (item1.DefectCauseId == null)
                                            {
                                                throw new SystemException("條碼:【" + item.BarcodeNo + "】必須處於回報完成狀態,才能做判定");

                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region //判斷是否相同人修改資料
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"SELECT JudgeUserId
                            //        FROM QMS.AqBarcode
                            //        WHERE AbnormalqualityId = @AbnormalqualityId
                            //        AND AqBarcodeId = @AqBarcodeId";
                            //dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                            //dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                            //result = sqlConnection.Query(sql, dynamicParameters);
                            //int? JudgeUserIdOriginal = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).JudgeUserId;

                            //if (JudgeUserIdOriginal != null)
                            //{
                            //    if (JudgeUserIdOriginal != CreateBy) throw new SystemException("請不要修改他人資料!!!");
                            //}
                            #endregion

                            if (QcType == "PVTQC")
                            {
                                if (JudgeStatus == "RW" || JudgeStatus == "SR") throw new SystemException("試產檢不能判定返修或是原站返修!!!");
                            }

                            #region //判斷該製程是否最後一站
                            if (QcType == "IPQC" || QcType == "NON")
                            {
                                int MaxSortNumber = -1;
                                int SortNumber = -1;
                                sql = @"SELECT c.MoId,b.MoProcessId,c.SortNumber,MAX(d1.SortNumber) MaxSortNumber
                                        FROM QMS.AqBarcode a
                                        INNER JOIN QMS.Abnormalquality b on a.AbnormalqualityId = b.AbnormalqualityId
                                        INNER JOIN MES.MoProcess c on b.MoProcessId = c.MoProcessId
                                        INNER JOIN MES.ManufactureOrder d on c.MoId = d.MoId
                                        INNER JOIN MES.MoProcess d1 on d.MoId = d1.MoId
								        WHERE a.AbnormalqualityId = @AbnormalqualityId
								        GROUP BY c.MoId,b.MoProcessId,c.SortNumber";
                                dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("資料有問題,請重新輸入");
                                foreach(var item in result)
                                {
                                    MaxSortNumber = item.MaxSortNumber;
                                    SortNumber = item.SortNumber;
                                }
                                if (MaxSortNumber == SortNumber)
                                {
                                    lastMoProcess = true;
                                }
                            }
                            #endregion

                            #region //判斷該製程是否委外製程
                            if (QcType != "PVTQC" && QcType != "IQC")
                            {
                                sql = @"SELECT TOP 1 a.Status
                                , b.BarcodeNo
                                , d.MoId, d.MoProcessId ,d.OspDetailId
                                FROM MES.OspReceiptBarcode a
                                INNER JOIN MES.OspBarcode b ON a.OspBarcodeId = b.OspBarcodeId
                                INNER JOIN MES.OspReceiptDetail c ON a.OsrDetailId = c.OsrDetailId
                                INNER JOIN MES.OspDetail d ON c.OspDetailId = d.OspDetailId
                                WHERE d.MoProcessId = @MoProcessId
                                AND b.BarcodeNo = @BarcodeNo
                                Order By c.CreateDate DESC";
                                dynamicParameters.Add("BarcodeNo", BarcodeNo);
                                if (QcType == "IPQC")
                                {
                                    dynamicParameters.Add("MoProcessId", MoProcessId);
                                }
                                else if (QcType == "OQC")
                                {
                                    dynamicParameters.Add("MoProcessId", CurrentMoProcessId);
                                }
                                else if (QcType == "NON")
                                {
                                    dynamicParameters.Add("MoProcessId", MoProcessId);
                                }

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0)
                                {
                                    OspMoProcess = true;
                                    OspDetailId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).OspDetailId;
                                }

                            }
                            #endregion

                            #region //判斷退回站別是否存在
                            if (JudgeReturnMoProcessId > 0)
                            {
                                dynamicParameters = new DynamicParameters();

                                sql = @"SELECT TOP 1 1
                                FROM MES.MoProcess a 
                                WHERE a.MoId = @MoId
                                AND a.MoProcessId = @JudgeReturnMoProcessId";
                                dynamicParameters.Add("MoId", MoId);
                                dynamicParameters.Add("JudgeReturnMoProcessId", JudgeReturnMoProcessId);

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("退回站別不存在，請重新輸入!");
                            }
                            #endregion

                            #region //撈取BarcodeProcessId
                            if (QcType != "PVTQC" && QcType != "IQC")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.BarcodeProcessId
                                FROM MES.BarcodeProcess a
                                WHERE a.MoProcessId = @MoProcessId
                                AND a.BarcodeId = @BarcodeId
                                AND a.StartDate = (SELECT MAX(z.StartDate)
                                                                FROM MES.BarcodeProcess z
                                                               WHERE z.BarcodeId = @BarcodeId
                                                                 AND z.MoProcessId = @MoProcessId)";
                                dynamicParameters.Add("BarcodeId", BarcodeId);
                                if (QcType == "IPQC")
                                {
                                    dynamicParameters.Add("MoProcessId", MoProcessId);
                                }
                                else if (QcType == "OQC")
                                {
                                    dynamicParameters.Add("MoProcessId", CurrentMoProcessId);
                                }
                                else if (QcType == "NON")
                                {
                                    dynamicParameters.Add("MoProcessId", MoProcessId);
                                }
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("資料有問題,請重新輸入");
                                BarcodeProcessId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).BarcodeProcessId;
                                insertResult.Add(JudgeStatus);
                                insertResult.Add(Convert.ToString(BarcodeProcessId));
                            }
                            #endregion

                            #region //更新判定資料
                            #region //更新Barcode NextMoProcessId
                            if(JudgeConfirm == "N")
                            {
                                //判定狀態為報廢時, 條碼流程為0, NextMoProcessId = -1
                                if (QcType != "IQC")
                                {
                                    switch (JudgeStatus)
                                    {
                                        #region //報廢
                                        case "S":
                                            #region 製程為托外時,修改MES.OspReceiptBarcode資料 托外條碼狀態
                                            if (OspMoProcess == true)
                                            {
                                                //修改 托外條碼狀態
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE a 
                                            SET
                                            a.Status = @JudgeStatus,
                                            a.LastModifiedBy = @LastModifiedBy,
                                            a.LastModifiedDate = @LastModifiedDate
                                            FROM MES.OspReceiptBarcode a
                                            INNER JOIN MES.OspBarcode b on a.OspBarcodeId = b.OspBarcodeId
                                            WHERE b.BarcodeNo = @BarcodeNo";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      JudgeStatus,
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      BarcodeNo
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                                //修改 MES.OspReceiptDetail資料 托外入庫詳細資料數量
                                                dynamicParameters = new DynamicParameters();

                                                sql = @"UPDATE MES.OspReceiptDetail SET
                                                        AvailableQty = AvailableQty - 1,
                                                        ReturnQty = ReturnQty + 1,
                                                        LastModifiedBy = @LastModifiedBy,
                                                        LastModifiedDate = @LastModifiedDate
                                                        WHERE OspDetailId = @OspDetailId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      OspDetailId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            }
                                            #endregion

                                            if (QcType != "PVTQC")
                                            {
                                                #region //修改MES.Barcode的資料 狀態+下一站
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE MES.Barcode SET
                                        BarcodeStatus = '0',
                                        NextMoProcessId = -1,
                                        CurrentProdStatus = @JudgeStatus,
                                        LastModifiedBy = @LastModifiedBy,
                                        LastModifiedDate = @LastModifiedDate
                                        WHERE BarcodeId = @BarcodeId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      JudgeStatus,
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      BarcodeId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                                #endregion

                                                #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                                                UpdateBarcodeProcessStatus("U", BarcodeProcessId, "S");
                                                #endregion
                                            }

                                            break;
                                        #endregion

                                        #region //返修
                                        case "RW":
                                            #region 製程為托外時,修改MES.OspReceiptBarcode資料 托外條碼狀態
                                            if (OspMoProcess == true)
                                            {
                                                //修改 托外條碼狀態
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE a 
                                            SET
                                            a.Status = @JudgeStatus,
                                            a.LastModifiedBy = @LastModifiedBy,
                                            a.LastModifiedDate = @LastModifiedDate
                                            FROM MES.OspReceiptBarcode a
                                            INNER JOIN MES.OspBarcode b on a.OspBarcodeId = b.OspBarcodeId
                                            WHERE b.OspBarcodeId = @BarcodeId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      JudgeStatus,
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      BarcodeId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                                //修改 托外入庫詳細資料數量
                                                dynamicParameters = new DynamicParameters();

                                                sql = @"UPDATE MES.OspReceiptDetail SET
                                                        AvailableQty = AvailableQty - 1,
                                                        ReturnQty = ReturnQty + 1,
                                                        LastModifiedBy = @LastModifiedBy,
                                                        LastModifiedDate = @LastModifiedDate
                                                        WHERE OspDetailId = @OspDetailId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      OspDetailId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            }
                                            #endregion

                                            #region //修改MES.Barcode的資料 狀態+下一站
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE MES.Barcode SET
                                        NextMoProcessId = @JudgeReturnMoProcessId,
                                        CurrentProdStatus = @JudgeStatus,
                                        BarcodeStatus='1',
                                        LastModifiedBy = @LastModifiedBy,
                                        LastModifiedDate = @LastModifiedDate
                                        WHERE BarcodeId = @BarcodeId";
                                            dynamicParameters.AddDynamicParams(
                                              new
                                              {
                                                  JudgeReturnMoProcessId,
                                                  JudgeStatus,
                                                  LastModifiedBy,
                                                  LastModifiedDate,
                                                  BarcodeId
                                              });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion

                                            #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                                            UpdateBarcodeProcessStatus("U", BarcodeProcessId, "F");
                                            #endregion
                                            break;
                                        #endregion

                                        #region //當站返修
                                        case "SR":
                                            #region 製程為托外時,修改MES.OspReceiptBarcode資料 托外條碼狀態
                                            if (OspMoProcess == true)
                                            {
                                                //修改 托外條碼狀態
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE a 
                                                        SET
                                                        a.Status = @JudgeStatus,
                                                        a.LastModifiedBy = @LastModifiedBy,
                                                        a.LastModifiedDate = @LastModifiedDate
                                                        FROM MES.OspReceiptBarcode a
                                                        INNER JOIN MES.OspBarcode b on a.OspBarcodeId = b.OspBarcodeId
                                                        WHERE b.OspBarcodeId = @BarcodeId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      JudgeStatus,
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      BarcodeId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                                //修改 MES.OspReceiptDetail資料 托外入庫詳細資料數量
                                                dynamicParameters = new DynamicParameters();

                                                sql = @"UPDATE MES.OspReceiptDetail SET
                                                        AvailableQty = AvailableQty - 1,
                                                        ReturnQty = ReturnQty + 1,
                                                        LastModifiedBy = @LastModifiedBy,
                                                        LastModifiedDate = @LastModifiedDate
                                                        WHERE OspDetailId = @OspDetailId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      OspDetailId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            }
                                            #endregion

                                            #region //修改MES.Barcode的資料 狀態+下一站
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE MES.Barcode SET
                                            BarcodeStatus='1',
                                            NextMoProcessId = CurrentMoProcessId,
                                            CurrentProdStatus = @JudgeStatus,
                                            LastModifiedBy = @LastModifiedBy,
                                            LastModifiedDate = @LastModifiedDate
                                            WHERE BarcodeId = @BarcodeId";
                                            dynamicParameters.AddDynamicParams(
                                              new
                                              {
                                                  JudgeStatus,
                                                  LastModifiedBy,
                                                  LastModifiedDate,
                                                  BarcodeId
                                              });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion

                                            #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                                            UpdateBarcodeProcessStatus("U", BarcodeProcessId, "F");
                                            #endregion
                                            break;
                                        #endregion

                                        #region //誤判 ,特採放行, 不良品/流程更換
                                        case "M":  //誤判
                                        case "R":  //特採放行
                                        case "NK": //不良品/流程更換
                                            #region //判斷是否最後一站
                                            if (QcType == "IPQC" || QcType == "NON")
                                            {
                                                if (lastMoProcess == true)
                                                {
                                                    BarcodeStatus = "0";
                                                }
                                            }
                                            #endregion

                                            #region //不良品/流程更換=>採用強制結束狀態
                                            if (JudgeStatus == "NK")
                                            {
                                                BarcodeStatus = "7";
                                            }
                                            #endregion

                                            #region //修改條碼狀態
                                            if (QcType != "PVTQC")
                                            {
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE MES.Barcode SET
                                                        CurrentProdStatus = @JudgeStatus,
                                                        BarcodeStatus = @BarcodeStatus,
                                                        LastModifiedBy = @LastModifiedBy,
                                                        LastModifiedDate = @LastModifiedDate
                                                        WHERE BarcodeId = @BarcodeId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      JudgeStatus,
                                                      BarcodeStatus,
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      BarcodeId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                                #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                                                UpdateBarcodeProcessStatus("U", BarcodeProcessId, "P");
                                                #endregion
                                            }
                                            #endregion

                                            break;
                                        #endregion

                                        default: throw new SystemException("判定結果未設定!");
                                    }
                                }
                                else
                                {
                                    #region //撈取進貨數量
                                    int ReceiptQty = 0;
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ReceiptQty
                                        FROM SCM.GrDetail a
                                        WHERE a.GrDetailId = @GrDetailId";
                                    dynamicParameters.Add("GrDetailId", GrDetailId);

                                    var resultReceiptQty = sqlConnection.Query(sql, dynamicParameters);
                                    foreach (var item in resultReceiptQty)
                                    {
                                        ReceiptQty = item.ReceiptQty;
                                    }
                                    if (ReceiptQty < ReleaseQty) throw new SystemException("特採數量不可超過進貨單進貨數" + ReceiptQty + "!!!");
                                    if (ReleaseQty < 0) throw new SystemException("特採數量不可為負!!!");
                                    #endregion

                                    switch (JudgeStatus)
                                    {
                                        #region //更新進貨單單身資料
                                        case "R":
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE SCM.GrDetail SET
                                                    QcStatus = @QcStatus,
                                                    AcceptQty = @ReleaseQty,
                                                    AvailableQty = @ReleaseQty,
                                                    ReturnQty = ReceiptQty - @ReleaseQty,
                                                    LastModifiedBy = @LastModifiedBy,
                                                    LastModifiedDate = @LastModifiedDate
                                                    WHERE GrDetailId = @GrDetailId";
                                            dynamicParameters.AddDynamicParams(
                                              new
                                              {
                                                  QcStatus = "4",
                                                  ReleaseQty = ReleaseQty,
                                                  LastModifiedBy,
                                                  LastModifiedDate,
                                                  GrDetailId
                                              });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            break;
                                        #endregion

                                        #region //更新進貨單單身資料
                                        case "S":
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE SCM.GrDetail SET
                                                    QcStatus = @QcStatus,
                                                    AcceptQty = @AcceptQty,
                                                    AvailableQty = @AvailableQty,
                                                    ReturnQty = ReceiptQty,
                                                    LastModifiedBy = @LastModifiedBy,
                                                    LastModifiedDate = @LastModifiedDate
                                                    WHERE GrDetailId = @GrDetailId";
                                            dynamicParameters.AddDynamicParams(
                                              new
                                              {
                                                  QcStatus = "3",
                                                  AcceptQty = 0,
                                                  AvailableQty = 0,
                                                  LastModifiedBy,
                                                  LastModifiedDate,
                                                  GrDetailId
                                              });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                            break;
                                        #endregion
                                        case "AM":
                                            break;
                                    }
                                }

                                if (ResponsibleUserId>0 && ResponsibleDeptId > 0)
                                {
                                    #region //更新品異單資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE QMS.AqBarcode SET
                                            JudgeStatus = @JudgeStatus,
                                            JudgeReturnMoProcessId = @JudgeReturnMoProcessId,
                                            JudgeReturnNextMoProcessId = @JudgeReturnNextMoProcessId,
                                            JudgeDesc = @JudgeDesc,
                                            JudgeUserId = @JudgeUserId,
                                            JudgeConfirm = @JudgeConfirm,
                                            JudgeDate = @JudgeDate,
                                            ResponsibleDeptId = @ResponsibleDeptId,
                                            ResponsibleUserId = @ResponsibleUserId,
                                            AqBarcodeStatus = @AqBarcodeStatus,
                                            ReleaseQty = @ReleaseQty,
                                            LastModifiedBy = @LastModifiedBy,
                                            LastModifiedDate = @LastModifiedDate
                                            WHERE AbnormalqualityId = @AbnormalqualityId
                                            AND AqBarcodeId = @AqBarcodeId";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          JudgeStatus,
                                          JudgeReturnMoProcessId = JudgeReturnMoProcessId > 0 ? JudgeReturnMoProcessId : nullData,
                                          JudgeReturnNextMoProcessId = JudgeReturnNextMoProcessId > 0 ? JudgeReturnNextMoProcessId : nullData,
                                          JudgeDesc,
                                          JudgeUserId,
                                          JudgeConfirm = 'Y',
                                          JudgeDate,
                                          ResponsibleDeptId,
                                          ResponsibleUserId,
                                          AqBarcodeStatus,
                                          ReleaseQty = QcType == "IQC"? ReleaseQty : (int?)null,
                                          LastModifiedBy,
                                          LastModifiedDate,
                                          AbnormalqualityId,
                                          AqBarcodeId
                                      });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    #region //更新品異單資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE QMS.AqBarcode SET
                                            JudgeStatus = @JudgeStatus,
                                            JudgeReturnMoProcessId = @JudgeReturnMoProcessId,
                                            JudgeReturnNextMoProcessId = @JudgeReturnNextMoProcessId,
                                            JudgeDesc = @JudgeDesc,
                                            JudgeUserId = @JudgeUserId,
                                            JudgeConfirm = @JudgeConfirm,
                                            JudgeDate = @JudgeDate,
                                            AqBarcodeStatus = @AqBarcodeStatus,
                                            ReleaseQty = @ReleaseQty,
                                            LastModifiedBy = @LastModifiedBy,
                                            LastModifiedDate = @LastModifiedDate
                                            WHERE AbnormalqualityId = @AbnormalqualityId
                                            AND AqBarcodeId = @AqBarcodeId";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          JudgeStatus,
                                          JudgeReturnMoProcessId = JudgeReturnMoProcessId > 0 ? JudgeReturnMoProcessId : nullData,
                                          JudgeReturnNextMoProcessId = JudgeReturnNextMoProcessId > 0 ? JudgeReturnNextMoProcessId : nullData,
                                          JudgeDesc,
                                          JudgeUserId,
                                          JudgeConfirm = 'Y',
                                          JudgeDate,
                                          AqBarcodeStatus,
                                          ReleaseQty = QcType == "IQC"? ReleaseQty : (int?)null,
                                          LastModifiedBy,
                                          LastModifiedDate,
                                          AbnormalqualityId,
                                          AqBarcodeId
                                      });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }

                            }
                            else
                            {
                                #region //更新品異單資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE QMS.AqBarcode SET
                                    ResponsibleUserId = @ResponsibleUserId,
                                    ResponsibleDeptId = @ResponsibleDeptId,
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE AbnormalqualityId = @AbnormalqualityId
                                    AND AqBarcodeId = @AqBarcodeId";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      ResponsibleUserId,
                                      ResponsibleDeptId,
                                      LastModifiedBy,
                                      LastModifiedDate,
                                      AbnormalqualityId,
                                      AqBarcodeId
                                  });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }


                            #endregion
                            #endregion
                        }

                        #region //判定原因是否納入常用片語
                        if (UserAqPhrase == "Y")
                        {
                            #region //判段該原因是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM QMS.AqPhrase a
                                    WHERE a.PhraseText = @PhraseText";
                            dynamicParameters.Add("PhraseText", JudgeDesc);

                            var resultDepartment = sqlConnection.Query(sql, dynamicParameters);
                            if (resultDepartment.Count() <= 0)
                            {
                                #region 常用片語建立
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO QMS.AqPhrase (CompanyId, PhraseText,
                                    CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @PhraseText,
                                    @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        CompanyId = CurrentCompany,
                                        PhraseText = JudgeDesc,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion
                        }
                        #endregion

                        #region //寄送品異單簽核完成信

                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
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

        #region //UpdateAqCountersignDetail -- 更新品異單異常會簽資料 -- Shintokuro 2022.10.17
        public string UpdateAqCountersignDetail(int AbnormalqualityId, int AqBarcodeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        string CheckOut = "N";

                        #region //判斷異常條碼判定資料是否正確
                        sql = @"SELECT ResponsibleDeptId, ResponsibleUserId, RepairCauseUserId, ResponsibleSupervisorId
                                , JudgeUserId, CountersignUserId
                                FROM QMS.AqBarcode
                                WHERE AbnormalqualityId = @AbnormalqualityId
                                AND AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("資料不存在,請重新確認");
                        foreach(var item in result)
                        {
                            if(item.ResponsibleDeptId <= 0) throw new SystemException("責任單位未維護,不可以做會簽,請重新確認");
                            if(item.ResponsibleUserId <= 0) throw new SystemException("責任者未維護,不可以做會簽,請重新確認");
                            if(item.RepairCauseUserId ==null) throw new SystemException("對策者未維護,不可以做會簽,請重新確認");
                            if(item.ResponsibleSupervisorId == null) throw new SystemException("責任主管未維護,不可以做會簽,請重新確認");
                            if(item.JudgeUserId == null) throw new SystemException("判定者未維護,不可以做會簽,請重新確認");
                            if(item.CountersignUserId > 0) throw new SystemException("已經完成會簽,請重新確認");
                        }
                        #endregion

                        #region //更新會簽主管資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.AqBarcode SET
                                CountersignUserId = @CountersignUserId,
                                AqBarcodeStatus = @AqBarcodeStatus,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE AbnormalqualityId = @AbnormalqualityId
                                AND AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              CountersignUserId = CreateBy,
                              AqBarcodeStatus = 5,
                              LastModifiedBy,
                              LastModifiedDate,
                              AbnormalqualityId,
                              AqBarcodeId
                          });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region /判斷單身是否結單
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT DISTINCT AqBarcodeStatus
                                FROM QMS.AqBarcode
                                WHERE AbnormalqualityId = @AbnormalqualityId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                        result = sqlConnection.Query(sql, dynamicParameters);

                        int failNum = 0;
                        int passNum = 0;

                        foreach (var item in result)
                        {
                            string aqBarcodeStatus = item.AqBarcodeStatus;
                            if (aqBarcodeStatus == "5")
                            {
                                passNum++;
                            }
                            else
                            {
                                failNum++;
                            }
                        }
                        #region //單頭結單
                        if (failNum == 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.Abnormalquality SET
                                    AbnormalqualityStatus = @AbnormalqualityStatus,
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE AbnormalqualityId = @AbnormalqualityId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  AbnormalqualityStatus = "P",
                                  LastModifiedBy,
                                  LastModifiedDate,
                                  AbnormalqualityId
                              });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            CheckOut = "Y";
                        }
                        #endregion
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "會簽成功",
                            data = CheckOut.Split(' ')

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

        #region //UpdateAqCountersignDetailRe -- 更新品異單異常判定資料會簽(撤銷) -- Shintokuro 2022.10.18
        public string UpdateAqCountersignDetailRe(int AbnormalqualityId, int AqBarcodeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        int? nullData = null;
                        string CheckOut = "N";

                        #region //判斷異常條碼判定資料是否正確
                        sql = @"SELECT CountersignUserId
                                FROM QMS.AqBarcode
                                WHERE AbnormalqualityId = @AbnormalqualityId
                                AND AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("資料不存在,請重新確認");
                        //int CountersignUserIdOriginal = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CountersignUserId;
                        //if (CountersignUserIdOriginal != CreateBy) throw new SystemException("請不要修改他人資料!!!");
                        #endregion

                        #region //更新撤銷會簽主管資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.AqBarcode SET
                                CountersignUserId = @CountersignUserId,
                                AqBarcodeStatus = @AqBarcodeStatus,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE AbnormalqualityId = @AbnormalqualityId
                                AND AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              CountersignUserId = nullData,
                              AqBarcodeStatus = 3,
                              LastModifiedBy,
                              LastModifiedDate,
                              AbnormalqualityId,
                              AqBarcodeId
                          });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region /判斷單身是否結單
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT DISTINCT AqBarcodeStatus
                                FROM QMS.AqBarcode
                                WHERE AbnormalqualityId = @AbnormalqualityId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                        result = sqlConnection.Query(sql, dynamicParameters);

                        int failNum = 0;
                        int passNum = 0;

                        foreach (var item in result)
                        {
                            string aqBarcodeStatus = item.AqBarcodeStatus;
                            if (aqBarcodeStatus == "5")
                            {
                                passNum++;
                            }
                            else
                            {
                                failNum++;
                            }
                        }
                        #region //單頭結單
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.Abnormalquality SET
                                AbnormalqualityStatus = @AbnormalqualityStatus,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE AbnormalqualityId = @AbnormalqualityId";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              AbnormalqualityStatus = failNum!=0 ? "F":"P",
                              LastModifiedBy,
                              LastModifiedDate,
                              AbnormalqualityId
                          });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "撤銷成功",
                            data = CheckOut.Split(' ')

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

        #region //UpdateAqCountersignDetailBatch -- 更新品異單異常判定資料會簽(批量) -- Shintokuro 2022.10.19
        public string UpdateAqCountersignDetailBatch(int AbnormalqualityId, string AqBarcodeIdList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        string CheckOut = "N";
                        string QcType = "";
                        foreach (var AqBarcodeId in AqBarcodeIdList.Split(','))
                        {
                            #region //判斷異常條碼判定資料是否正確
                            sql = @"SELECT a.ResponsibleDeptId, a.ResponsibleUserId, a.RepairCauseUserId, a.ResponsibleSupervisorId
                                    , a.JudgeUserId, a.CountersignUserId,a1.QcType
                                    , b.BarcodeNo
                                    FROM QMS.AqBarcode a
                                    INNER JOIN QMS.Abnormalquality a1 on a.AbnormalqualityId = a1.AbnormalqualityId
                                    LEFT JOIN MES.Barcode b on a.BarcodeId = b.BarcodeId
                                    WHERE a.AbnormalqualityId = @AbnormalqualityId
                                    AND a.AqBarcodeId = @AqBarcodeId";
                            dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                            dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("資料不存在,請重新確認");
                            foreach (var item in result)
                            {
                                QcType = item.QcType;
                                if(QcType != "IQC")
                                {
                                    if (item.ResponsibleDeptId <= 0) throw new SystemException("條碼:【" + item.BarcodeNo + "】責任單位未維護,不可以做會簽,請重新確認");
                                    if (item.ResponsibleUserId <= 0) throw new SystemException("條碼:【" + item.BarcodeNo + "】責任者未維護,不可以做會簽,請重新確認");
                                    if (item.RepairCauseUserId == null) throw new SystemException("條碼:【" + item.BarcodeNo + "】對策者未維護,不可以做會簽,請重新確認");
                                    if (item.ResponsibleSupervisorId == null) throw new SystemException("條碼:【" + item.BarcodeNo + "】責任主管未維護,不可以做會簽,請重新確認");
                                    if (item.JudgeUserId == null) throw new SystemException("條碼:【" + item.BarcodeNo + "】判定者未維護,不可以做會簽,請重新確認");
                                    if (item.CountersignUserId > 0) throw new SystemException("條碼:【" + item.BarcodeNo + "】已經完成會簽,請重新確認");
                                }
                                else
                                {
                                    if (item.ResponsibleDeptId <= 0) throw new SystemException("責任單位未維護,不可以做會簽,請重新確認");
                                    if (item.ResponsibleUserId <= 0) throw new SystemException("責任者未維護,不可以做會簽,請重新確認");
                                    if (item.RepairCauseUserId == null) throw new SystemException("對策者未維護,不可以做會簽,請重新確認");
                                    if (item.ResponsibleSupervisorId == null) throw new SystemException("責任主管未維護,不可以做會簽,請重新確認");
                                    if (item.JudgeUserId == null) throw new SystemException("判定者未維護,不可以做會簽,請重新確認");
                                    if (item.CountersignUserId > 0) throw new SystemException("已經完成會簽,請重新確認");
                                }
                            }
                            #endregion

                            #region //更新會簽主管資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.AqBarcode SET
                                    CountersignUserId = @CountersignUserId,
                                    AqBarcodeStatus = @AqBarcodeStatus,
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE AbnormalqualityId = @AbnormalqualityId
                                    AND AqBarcodeId = @AqBarcodeId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  CountersignUserId = CreateBy,
                                  AqBarcodeStatus = 5,
                                  LastModifiedBy,
                                  LastModifiedDate,
                                  AbnormalqualityId,
                                  AqBarcodeId
                              });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

                        #region /判斷單身是否結單
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT DISTINCT AqBarcodeStatus
                                FROM QMS.AqBarcode
                                WHERE AbnormalqualityId = @AbnormalqualityId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                        var result1 = sqlConnection.Query(sql, dynamicParameters);

                        string AbnormalqualityStatus = "P";
                        int failNum = 0;
                        int passNum = 0;

                        foreach (var item in result1)
                        {
                            string aqBarcodeStatus = item.AqBarcodeStatus;
                            if (aqBarcodeStatus == "5")
                            {
                                passNum++;
                            }
                            else
                            {
                                failNum++;
                            }
                        }
                        #region //單頭結單
                        if (failNum == 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.Abnormalquality SET
                                    AbnormalqualityStatus = @AbnormalqualityStatus,
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE AbnormalqualityId = @AbnormalqualityId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  AbnormalqualityStatus,
                                  LastModifiedBy,
                                  LastModifiedDate,
                                  AbnormalqualityId
                              });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            CheckOut = "Y";
                        }
                        #endregion
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "會簽成功",
                            data = CheckOut.Split(' ')

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

        #region //Change BarcodeProcess Status & Update MoProcess Total Quantity
        public void UpdateBarcodeProcessStatus(string TxType, int BarcodeProcessId, string ProdStatus)
        {
            if (ProdStatus == "") throw new Exception("ProdStatus 不可為空");
            if (BarcodeProcessId == -1) throw new Exception("BarcodeProcessId 不可為空");
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region default
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        int PassQty = 0;
                        int NgQty = 0;
                        int ScrapQty = 0;
                        int PreMoProcessId = -1;
                        #endregion

                        #region //找出BarcodeProcess資訊

                        sql = @"SELECT MoProcessId,StationQty,ProdStatus,BarcodeId
                                  FROM MES.BarcodeProcess 
                                 WHERE BarcodeProcessId = @BarcodePRocessId";

                        dynamicParameters.Add("BarcodePRocessId", BarcodeProcessId);
                        var BarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);
                        if (BarcodeProcessResult.Count() > 0)
                        {
                            foreach (var item in BarcodeProcessResult)
                            {
                                int MoProcessId = Convert.ToInt32(item.MoProcessId);
                                int StationQty = Convert.ToInt32(item.StationQty);
                                string OrigProdStatus = item.ProdStatus.ToString();
                                int BarcodeId = Convert.ToInt32(item.BarcodeId);
                                int TransferQty = 0;

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT y.MoProcessId
                                          FROM MES.Barcode x
                                               INNER JOIN MES.BarcodeProcess y ON x.BarcodeId = y.BarcodeId
                                         WHERE x.BarcodeId = @BarcodeId
                                           AND y.StartDate = (SELECT MAX(z.StartDate)
                                                                FROM MES.BarcodeProcess z
					                                           WHERE z.BarcodeId = x.BarcodeId
					                                             AND z.BarcodeProcessId != @BarcodeProcessId)";
                                dynamicParameters.Add("BarcodeId", BarcodeId);
                                dynamicParameters.Add("BarcodePRocessId", BarcodeProcessId);
                                var PreMoProcessResult = sqlConnection.Query(sql, dynamicParameters);
                                if (PreMoProcessResult.Count() > 0)
                                {
                                    PreMoProcessId = Convert.ToInt32(sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MoProcessId);
                                }

                                if (ProdStatus.Equals("S"))
                                {
                                    ScrapQty = StationQty;
                                    PassQty = 0;
                                    NgQty = 0;
                                }
                                else if (ProdStatus.Equals("F"))
                                {
                                    ScrapQty = 0;
                                    PassQty = 0;
                                    NgQty = StationQty;
                                }
                                else if (ProdStatus.Equals("P"))
                                {
                                    ScrapQty = 0;
                                    PassQty = StationQty;
                                    NgQty = 0;
                                }
                                else
                                {
                                    throw new SystemException("條碼狀態錯誤!");
                                }

                                switch (TxType)
                                {
                                    case "A": //過站時需Update MoProcess , 加上新狀態數量
                                        #region //取出這條碼有沒有NG或報廢數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT SUM(TransactionQty) TransactionQty
                                                  FROM MES.BarcodeTransfer a
                                                 WHERE a.FromBarcodeId = @BarcodeId
                                                   AND EXISTS(SELECT 1
                                                                FROM MES.BarcodeProcess b
			                                                   WHERE a.ToBarcodeId = b.BarcodeId
			                                                     AND b.MoProcessId = @MoProcessId)";
                                        dynamicParameters.Add("BarcodeId", BarcodeId);
                                        dynamicParameters.Add("MoProcessId", MoProcessId);
                                        var BarcodeNgQtyResult = sqlConnection.Query(sql, dynamicParameters);
                                        if (BarcodeNgQtyResult.Count() > 0)
                                        {
                                            TransferQty = Convert.ToInt32(sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TransactionQty);
                                        }
                                        #endregion

                                        #region //更新MoProcess Quantity Informations
                                        sql = @"UPDATE MES.MoProcess 
                                                   SET TotalInputQty = TotalInputQty + @StationQty,
                                                       TotalPassQty = TotalPassQty + @PassQty + @TransferQty,
                                                       TotalNgQty = TotalNgQty + @NgQty,
                                                       TotalScrapQty = TotalScrapQty + @ScrapQty
                                                 WHERE MoProcessId = @MoProcessId";
                                        dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            StationQty,
                                            PassQty,
                                            TransferQty,
                                            NgQty,
                                            ScrapQty,
                                            MoProcessId
                                        });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //更新WipQty
                                        if (!PreMoProcessId.Equals(MoProcessId))
                                        {
                                            sql = @"UPDATE MES.MoProcess 
                                                   SET WipQty = WipQty - @StationQty - @TransferQty
                                                 WHERE MoProcessId = @PreMoProcessId";
                                            dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                StationQty,
                                                TransferQty,
                                                PreMoProcessId
                                            });

                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                            sql = @"UPDATE MES.MoProcess 
                                                   SET WipQty = WipQty + @StationQty + @TransferQty
                                                 WHERE MoProcessId = @MoProcessId";
                                            dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                StationQty,
                                                TransferQty,
                                                MoProcessId
                                            });

                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        }
                                        #endregion
                                        break;
                                    case "U": //過站時Update MoProcess , 扣掉原本狀態數量, 加上新狀態數量
                                        #region //更新BarcodeProcess ProdStatus
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE MES.BarcodeProcess 
                                                   SET ProdStatus = @ProdStatus
                                                 WHERE BarcodeProcessId = @BarcodeProcessId";
                                        dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            ProdStatus,
                                            BarcodeProcessId
                                        });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //更新MoProcess Quantity Informations
                                        if (!OrigProdStatus.Equals(ProdStatus))
                                        {
                                            sql = "UPDATE MES.MoProcess ";
                                            #region //判斷原始BarcodeProcess.ProdStatus & New ProdStatus
                                            if (OrigProdStatus.Equals("P") && ProdStatus.Equals("F"))
                                            {
                                                //扣掉PassQty, 增加NgQty
                                                sql += "SET TotalPassQty = TotalPassQty - @StationQty,";
                                                sql += " TotalNgQty = TotalNgQty + @StationQty ";
                                            }
                                            else if (OrigProdStatus.Equals("P") && ProdStatus.Equals("S"))
                                            {
                                                //扣掉PassQty, 增加ScrapQty
                                                sql += "SET TotalPassQty = TotalPassQty - @StationQty,";
                                                sql += " TotalScrapQty = TotalScrapQty + @StationQty ";
                                            }
                                            else if (OrigProdStatus.Equals("F") && ProdStatus.Equals("P"))
                                            {
                                                //扣掉NgQty, 增加PassQty
                                                sql += "SET TotalNgQty = TotalNgQty - @StationQty,";
                                                sql += " TotalPassQty = TotalPassQty + @StationQty ";
                                            }
                                            else if (OrigProdStatus.Equals("F") && ProdStatus.Equals("S"))
                                            {
                                                //扣掉NgQty, 增加ScripQty
                                                sql += "SET TotalNgQty = TotalNgQty - @StationQty,";
                                                sql += " TotalScrapQty = TotalScrapQty + @StationQty ";
                                            }
                                            else if (OrigProdStatus.Equals("S") && ProdStatus.Equals("F"))
                                            {
                                                //扣掉ScripQty, 增加NgQty
                                                sql += "SET TotalScrapQty = TotalScrapQty - @StationQty,";
                                                sql += " TotalNgQty = TotalNgQty + @StationQty ";
                                            }
                                            else if (OrigProdStatus.Equals("S") && ProdStatus.Equals("P"))
                                            {
                                                //扣掉ScripQty, 增加PassQty
                                                sql += "SET TotalScrapQty = TotalScrapQty - @StationQty,";
                                                sql += " TotalPassQty = TotalPassQty + @StationQty ";
                                            }
                                            else
                                            {
                                                throw new SystemException("UpdateBarcodeProcessStatus非預期錯誤");
                                            }
                                            #endregion
                                            sql += " WHERE MoProcessId = @MoProcessId ";
                                            dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                StationQty,
                                                MoProcessId
                                            });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                        }
                                        #endregion
                                        break;
                                    case "D": //過站時Update MoProcess , 扣掉原本狀態數量
                                        #region //更新MoProcess Quantity Informations
                                        if (OrigProdStatus.Equals("P"))
                                        {
                                            //扣掉PassQty
                                        }
                                        else if (OrigProdStatus.Equals("F"))
                                        {
                                            //扣掉NgQty
                                        }
                                        else if (OrigProdStatus.Equals("S"))
                                        {
                                            //扣掉ScrapQty
                                        }
                                        #endregion
                                        break;
                                    default: throw new SystemException("TxType 未定義");
                                }

                            }
                        }
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

        }

        #endregion

        #endregion

        #region //Delete

        #region //DeleteAbnormalquality -- 刪除品異單 -- Shintokuro 2022-10-21
        public string DeleteAbnormalquality(int AbnormalqualityId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                    
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        List<int> fileIdList = new List<int>();
                        List<int> aqQcItemIdList = new List<int>();
                        List<int> barcodeIdList = new List<int>();
                        int rowsAffected = 0;

                        #region //判斷品異單資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.Abnormalquality
                                WHERE AbnormalqualityId = @AbnormalqualityId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品異單單頭資料錯誤!");
                        #endregion

                        #region //判斷品異單單身是否有資料 進階
                        sql = @"SELECT b1.CurrentProdStatus,a.QcType,b.BarcodeId
                                FROM QMS.Abnormalquality a
                                INNER JOIN QMS.AqBarcode b on a.AbnormalqualityId = b.AbnormalqualityId
                                INNER JOIN MES.Barcode b1 on b.BarcodeId = b1.BarcodeId
                                WHERE b.AbnormalqualityId = @AbnormalqualityId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        //if (result.Count() > 0) throw new SystemException("該品異單已經有異常條碼資料,不能刪除");
                        foreach(var item in result)
                        {
                            if(item.QcType == "NON") throw new SystemException("異常條碼經由進階功能判斷為不良品,不能刪除");
                            //if(item.CurrentProdStatus == "F") throw new SystemException("異常條碼經由進階功能判斷為不良品,不能刪除");
                            barcodeIdList.Add(item.BarcodeId);
                        }
                        #endregion



                        #region //判斷是不是有做過工程檢的入庫單
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT Count(a.QcBarcodeId) haveQcRecordNum
                                FROM QMS.AqBarcode a
                                INNER JOIN QMS.Abnormalquality b on a.AbnormalqualityId = b.AbnormalqualityId
                                WHERE b.AbnormalqualityId = @AbnormalqualityId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

                        result = sqlConnection.Query(sql, dynamicParameters);

                        int haveQcRecordNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).haveQcRecordNum;

                        if (haveQcRecordNum > 0) throw new SystemException("該品異單經由工程檢開立,不能刪除!!!");
                        #endregion

                        #region //判斷異常條碼有沒有經主管對策確認
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 1
                                FROM QMS.AqBarcode a
                                INNER JOIN QMS.Abnormalquality b on a.AbnormalqualityId = b.AbnormalqualityId
                                WHERE b.AbnormalqualityId = @AbnormalqualityId
                                AND a.JudgeConfirm = 'Y'";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("品異單身已經有做判定,不能刪除!!!");
                        #endregion

                        #region //撈取佐證檔案ID
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT FileId
                                FROM QMS.AqFile a
                                INNER JOIN QMS.AqBarcode b on a.AqBarcodeId = b.AqBarcodeId
                                INNER JOIN QMS.Abnormalquality c on b.AbnormalqualityId = c.AbnormalqualityId
                                WHERE c.AbnormalqualityId = @AbnormalqualityId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

                        result = sqlConnection.Query(sql, dynamicParameters);

                        if(result.Count() > 0)
                        {
                            foreach(var item in result)
                            {
                                fileIdList.Add(item.FileId);
                            }
                        }
                        #endregion

                        #region //撈取多筆型態的異常項目對策資料
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT AqQcItemId
                                FROM QMS.AqQcItem a
                                INNER JOIN QMS.AqBarcode b on a.AqBarcodeId = b.AqBarcodeId
                                INNER JOIN QMS.Abnormalquality c on b.AbnormalqualityId = c.AbnormalqualityId
                                WHERE c.AbnormalqualityId = @AbnormalqualityId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

                        result = sqlConnection.Query(sql, dynamicParameters);

                        if (result.Count() > 0)
                        {
                            foreach (var item in result)
                            {
                                aqQcItemIdList.Add(item.AqQcItemId);
                            }
                        }
                        #endregion

                        #region //先刪除關聯Table
                        //品異單 - 單身 - 佐證檔案
                        if (fileIdList.Count() > 0)
                        {
                            //刪除QMS的紀錄 AqFile
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a FROM QMS.AqFile a
                                INNER JOIN QMS.AqBarcode b on a.AqBarcodeId = b.AqBarcodeId
                                WHERE b.AbnormalqualityId = @AbnormalqualityId";
                            dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                            //刪除BAS的檔案
                            foreach (var FileId in fileIdList)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a FROM BAS.[File] a
                                WHERE a.FileId = @FileId";
                                dynamicParameters.Add("FileId", FileId);
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            }
                        }

                        if (aqQcItemIdList.Count() > 0)
                        {
                            //刪除QcItem的紀錄
                            
                            foreach (var AqQcItemId in aqQcItemIdList)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a FROM QMS.AqQcItem a
                                WHERE a.AqQcItemId = @AqQcItemId";
                                dynamicParameters.Add("AqQcItemId", AqQcItemId);
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            }
                        }

                        //品異單 - 單身(異常條碼)
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM QMS.AqBarcode
                                WHERE AbnormalqualityId = @AbnormalqualityId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        //品異單 - 單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.Abnormalquality
                                WHERE AbnormalqualityId = @AbnormalqualityId";
                        dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        if(barcodeIdList.Count()> 0)
                        {
                            foreach(var barcodeId in barcodeIdList)
                            {
                                #region //判斷異常條碼是否為不良品流程
                                string BarcodeStatus = "";
                                string CurrentProdStatus = "P";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.BarcodeStatus
                                        FROM MES.Barcode a
                                        WHERE a.BarcodeId = @BarcodeId";
                                dynamicParameters.Add("BarcodeId", barcodeId);

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0)
                                {
                                    foreach (var item in result)
                                    {
                                        BarcodeStatus = item.BarcodeStatus;
                                        if (BarcodeStatus == "7")
                                        {
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 a.JudgeStatus
                                                FROM QMS.AqBarcode a
                                                WHERE a.BarcodeId = @BarcodeId
                                                ORDER BY a.AqBarcodeId DESC";
                                            dynamicParameters.Add("BarcodeId", barcodeId);

                                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                                            if (result.Count() > 0)
                                            {
                                                foreach (var item1 in result1)
                                                {
                                                    CurrentProdStatus = item1.JudgeStatus;
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion

                                #region //更改異常條碼目前狀態
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.Barcode SET
                                        CurrentProdStatus = @CurrentProdStatus,
                                        LastModifiedBy = @LastModifiedBy,
                                        LastModifiedDate = @LastModifiedDate
                                        WHERE BarcodeId = @barcodeId";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      CurrentProdStatus,
                                      LastModifiedBy,
                                      LastModifiedDate,
                                      barcodeId
                                  });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                        }

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

        #region //DeleteAqDetail -- 刪除品異單單身(異常條碼) -- Shintokuro 2022-10-21
        public string DeleteAqDetail(int AqBarcodeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                    
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        List<int> fileIdList = new List<int>();
                        int rowsAffected = 0;

                        #region //判斷品異單單身資料是否正確
                        string QcType = "";
                        int ResponsibleUserId = -1;
                        int? BarcodeIdBase = -1;
                        sql = @"SELECT a.ResponsibleUserId,a.BarcodeId,b.QcType
                                FROM QMS.AqBarcode a
                                INNER JOIN QMS.Abnormalquality b on a.AbnormalqualityId = b.AbnormalqualityId
                                WHERE a.AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品異單單身(異常條碼)資料錯誤!");
                        foreach(var item in result)
                        {
                            QcType = item.QcType;
                            ResponsibleUserId = item.ResponsibleUserId;
                            BarcodeIdBase = item.BarcodeId;
                        }
                        if (QcType == "NON") throw new SystemException("異常條碼經由進階功能判斷為不良品,不能刪除");

                        #endregion

                        #region //判斷是不是有做過工程檢的入庫單
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT count(a.QcBarcodeId) haveQcRecord
                                FROM QMS.AqBarcode a
                                WHERE a.AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        result = sqlConnection.Query(sql, dynamicParameters);

                        int? havecRecord = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).havecRecord;

                        if (havecRecord > 0) throw new SystemException("該異常條碼經由工程檢開立,不能刪除!!!");
                        #endregion

                        #region //判斷異常條碼有沒有經主管對策確認
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT Count(a.AqBarcodeStatus) haveResponsibleSupervisorId
                                FROM QMS.AqBarcode a
                                WHERE a.AqBarcodeId = @AqBarcodeId
                                AND a.AqBarcodeStatus in (3,4)";
                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        result = sqlConnection.Query(sql, dynamicParameters);

                        int haveResponsibleSupervisorId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).haveResponsibleSupervisorId;
                        if (haveResponsibleSupervisorId > 0) throw new SystemException("已經有主管對異常條碼做對策確認,不能刪除!!!");
                        #endregion

                        #region //撈取佐證檔案ID
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT FileId
                                FROM QMS.AqFile a
                                INNER JOIN QMS.AqBarcode b on a.AqBarcodeId = b.AqBarcodeId
                                WHERE b.AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        result = sqlConnection.Query(sql, dynamicParameters);

                        if (result.Count() > 0)
                        {
                            foreach (var item in result)
                            {
                                fileIdList.Add(item.FileId);
                            }
                        }
                        #endregion

                        #region //先刪除關聯Table
                        //品異單 - 單身 - 佐證檔案
                        if (fileIdList.Count() > 0)
                        {
                            //刪除QMS的紀錄
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a FROM QMS.AqFile a
                                WHERE a.AqBarcodeId = @AqBarcodeId";
                            dynamicParameters.Add("AqBarcodeId", AqBarcodeId);
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                            //刪除BAS的檔案
                            foreach (var FileId in fileIdList)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a FROM BAS.[File] a
                                WHERE a.FileId = @FileId";
                                dynamicParameters.Add("FileId", FileId);
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            }
                        }
                        #endregion

                        #region //刪除主要table
                        //品異單 - 單身(異常條碼)
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM QMS.AqBarcode
                                WHERE AqBarcodeId = @AqBarcodeId";
                        dynamicParameters.Add("AqBarcodeId", AqBarcodeId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更改異常條碼目前狀態
                        if(QcType != "IQC")
                        {
                            #region //判斷異常條碼是否為不良品流程
                            string BarcodeStatus = "";
                            string CurrentProdStatus = "P";
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.BarcodeStatus
                                    FROM MES.Barcode a
                                    WHERE a.BarcodeId = @BarcodeId";
                            dynamicParameters.Add("BarcodeId", BarcodeIdBase);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                foreach (var item in result)
                                {
                                    BarcodeStatus = item.BarcodeStatus;
                                    if (BarcodeStatus == "7")
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 JudgeStatus
                                                FROM QMS.AqBarcode a
                                                WHERE a.BarcodeId = @BarcodeId
                                                ORDER BY a.AqBarcodeId DESC";
                                        dynamicParameters.Add("BarcodeId", BarcodeIdBase);

                                        var result1 = sqlConnection.Query(sql, dynamicParameters);
                                        if (result.Count() > 0)
                                        {
                                            foreach (var item1 in result1)
                                            {
                                                CurrentProdStatus = item1.JudgeStatus;
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion


                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.Barcode SET
                                    CurrentProdStatus = @CurrentProdStatus,
                                    LastModifiedBy = @LastModifiedBy,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE BarcodeId = @BarcodeIdBase";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  CurrentProdStatus,
                                  LastModifiedBy,
                                  LastModifiedDate,
                                  BarcodeIdBase
                              });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        }
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

        #region //DeleteAqFileId -- 品異條碼原因佐證檔案 -- Shintokuro 2022-10-14
        public string DeleteAqFileId(int AqFileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                    
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷品異條碼資料是否正確
                        sql = @"SELECT TOP 1 FileId
                                FROM QMS.AqFile
                                WHERE AqFileId = @AqFileId";
                        dynamicParameters.Add("AqFileId", AqFileId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品異條碼資料錯誤!");
                        int FileId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).FileId;
                        #endregion

                        #region //判斷檔案是否存在
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 FileId
                                FROM BAS.[File]
                                WHERE FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("檔案不存在,無法執行刪除!");
                        #endregion

                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM QMS.AqFile
                                WHERE AqFileId = @AqFileId";
                        dynamicParameters.Add("AqFileId", AqFileId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM BAS.[File]
                                WHERE FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #endregion

        #region 品異單寄送
        private void SendIQcAqConfirmMail(SqlConnection sqlConnection, int AbnormalqualityId)
        {
            #region //自動送信件

            #region //取得MailServer
            dynamicParameters = new DynamicParameters();
            sql = @"select *
                    from BAS.MailServer
                    where CompanyId = @CompanyId";
            dynamicParameters.Add("CompanyId", CurrentCompany);

            var resultMailServer = sqlConnection.Query(sql, dynamicParameters);
            if (resultMailServer.Count() <= 0) throw new SystemException("MailServer設定錯誤!");
            var host = "";
            var account = "";
            var password = "";
            var port = 0;
            var sendMode = 0;

            foreach (var item in resultMailServer)
            {
                host = item.Host;
                account = item.Account;
                password = item.Password;
                port = Convert.ToInt32(item.Port);
                sendMode = Convert.ToInt32(item.SendMode);

            }
            #endregion

            #region //取得所有採購部門人員名單與信箱
            dynamicParameters = new DynamicParameters();
            sql = @"select b.DepartmentName, a.Job, a.UserName, a.Email 
                    from BAS.[User] a
                    inner join BAS.Department b on a.DepartmentId = b.DepartmentId
                    where a.UserStatus = 'F' and a.SystemStatus = 'A' and b.DepartmentId = @DepartmentId";

            dynamicParameters.Add("DepartmentId", 72);

            var resultDepartment = sqlConnection.Query(sql, dynamicParameters);
            if (resultDepartment.Count() <= 0) throw new SystemException("部門設定錯誤!");
            string departmentMailString = string.Join(";", resultDepartment.Select(user =>
                                    $"{user.DepartmentName}{(string.IsNullOrEmpty(user.Job) ? "" : "-" + user.Job)}-{user.UserName}:{user.Email}"));
            #endregion

            #region //取得所有品保部門主管人員名單與信箱
            dynamicParameters = new DynamicParameters();
            sql = @"select b.DepartmentName, a.Job, a.UserName, a.Email , a.JobType
                    from BAS.[User] a
                    inner join BAS.Department b on a.DepartmentId = b.DepartmentId
                    where a.JobType = '管理制' and (a.DepartmentId = 1519 OR a.DepartmentId = 1520) and a.UserStatus = 'F'";

            var resultQcDepartment = sqlConnection.Query(sql, dynamicParameters);
            if (resultQcDepartment.Count() <= 0) throw new SystemException("部門設定錯誤!");
            string resultQcDepartmentstr = string.Join(";", resultQcDepartment.Select(user =>
                                    $"{user.DepartmentName}{(string.IsNullOrEmpty(user.Job) ? "" : "-" + user.Job)}-{user.UserName}:{user.Email}"));
            #endregion

            #region //取得所有倉庫部門人員名單與信箱
            dynamicParameters = new DynamicParameters();
            sql = @"select b.DepartmentName, a.Job, a.UserName, a.Email 
                    from BAS.[User] a
                    inner join BAS.Department b on a.DepartmentId = b.DepartmentId
                    where a.UserStatus = 'F' and a.SystemStatus = 'A' and b.DepartmentId = @DepartmentId";

            dynamicParameters.Add("DepartmentId", 77);

            var resultMweDepartment = sqlConnection.Query(sql, dynamicParameters);
            if (resultMweDepartment.Count() <= 0) throw new SystemException("部門設定錯誤!");
            string mweDepartmentMailString = string.Join(";", resultMweDepartment.Select(user =>
                                    $"{user.DepartmentName}{(string.IsNullOrEmpty(user.Job) ? "" : "-" + user.Job)}-{user.UserName}:{user.Email}"));
            #endregion


            #region //取得信件內容相關參數
            dynamicParameters = new DynamicParameters();
            sql = @"select a.AbnormalqualityNo, (c.GrErpPrefix + '-' + c.GrErpNo + '-' + b.GrSeq) GrErpFullNo
                    , d.MtlItemNo,d.MtlItemName
                    from QMS.Abnormalquality a
                    inner join SCM.GrDetail b on a.GrDetailId = b.GrDetailId
                    inner join SCM.GoodsReceipt c on b.GrId = c.GrId
                    inner join PDM.MtlItem d on b.MtlItemId = d.MtlItemId
                    where a.AbnormalqualityId = @AbnormalqualityId";
            dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

            var resultQcRecord = sqlConnection.Query(sql, dynamicParameters);
            if (resultQcRecord.Count() <= 0) throw new SystemException("量測單據錯誤!");
            var GrErpFullNo = "";
            var AbnormalqualityNo = "";
            var MtlItemNo = "";
            var MtlItemName = "";

            foreach (var item in resultQcRecord)
            {
                GrErpFullNo = item.GrErpFullNo;
                AbnormalqualityNo = item.AbnormalqualityNo;
                MtlItemNo = item.MtlItemNo;
                MtlItemName = item.MtlItemName;
            }
            #endregion

            string mailToString = mweDepartmentMailString ;
            string mailCCString = departmentMailString + ";" + resultQcDepartmentstr + ";總經理-許智程:tim@zy-tech.com.tw;系統開發室-李宗宸:joe_li@zy-tech.com.tw;系統開發室-丁培文:pawun_ding@zy-tech.com.tw";


            #region 內文
            string mailSubject = "進貨檢驗品異單處理完成通知";

            string mailContent = $@"
                                        各位倉儲同仁您好：<br><br>

                                        本次進貨檢之品異判定作業已完成，請倉庫依照判定結果進行後續貨物查收或處理作業。<br><br>
                                        如需協助，請與採購部門及品保部門聯繫。<br><br>
                                        
                                        相關資訊如下：<br>
                                        ‧ 品異單號：{AbnormalqualityNo}<br>
                                        ‧ 進貨單號：{GrErpFullNo}<br>
                                        ‧ 品號：{MtlItemNo}<br>
                                        ‧ 品名：{MtlItemName}<br>
                                        
                                        敬祝<br>
                                        工作順利<br>
                                        ";
            #endregion

            #endregion

            #region //設定附檔
            string QcFileNameCondition = "";
            List<string> FilePath = new List<string>();
            List<string> NewFilePath = new List<string>();
            List<MailFile> mailFiles = new List<MailFile>();
            #endregion

            #region //寄送Mail
            MailConfig mailConfig = new MailConfig
            {
                Host = host,
                Port = port,
                SendMode = sendMode,
                From = "企業管理平台:jmo-service@zy-tech.com.tw",
                Subject = mailSubject,
                Account = account,
                Password = password,
                MailTo = mailToString,
                MailCc = "總經理-許智程:tim@zy-tech.com.tw;系統開發室-李宗宸:joe_li@zy-tech.com.tw;系統開發室-丁培文:pawun_ding@zy-tech.com.tw",
                //MailTo = "系統開發室-李宗宸:joe_li@zy-tech.com.tw;系統開發室-丁培文:pawun_ding@zy-tech.com.tw",
                //MailCc = "系統開發室-李宗宸:joe_li@zy-tech.com.tw;系統開發室-丁培文:pawun_ding@zy-tech.com.tw",

                MailBcc = "",
                HtmlBody = mailContent,
                TextBody = "-",
                FileInfo = mailFiles,
                QcFileFlag = "Y"
            };
            BaseHelper.MailSend(mailConfig);
            #endregion
        }
        #endregion



    }
}
