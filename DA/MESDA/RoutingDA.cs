using Dapper;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Transactions;
using System.Web;

namespace MESDA
{
    public class RoutingDA
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
        public SqlQuery sqlQuery = new SqlQuery();
        public DynamicParameters dynamicParameters = new DynamicParameters();

        public RoutingDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];

            GetUserInfo();

            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;
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
        #region //GetRouting -- 取得途程資料 -- Ann 2022-07-14
        public string GetRouting(int RoutingId, int ModeId, string RoutingType, string RoutingName, int RoutingItemId, string MtlItemNo, string MtlItemName
            , string Status, string StartDate, string EndDate, int UserId,string ProcessIds
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.RoutingId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ModeId, a.RoutingType, a.RoutingName, a.[Status], a.RoutingConfirm
                        , b.ModeNo, b.ModeName
                        , c.TypeNo RoutingType, c.TypeName RountingTypeName
                        , d.StatusNo, d.StatusName
						, (e.UserNo + '-' + e.UserName) UserName";
                    sqlQuery.mainTables =
                        @"FROM MES.Routing a
                        INNER JOIN MES.ProdMode b ON a.ModeId = b.ModeId
                        INNER JOIN BAS.[Type] c ON a.RoutingType = c.TypeNo AND c.TypeSchema = 'Routing.RoutingType'
                        INNER JOIN BAS.[Status] d ON a.[Status] = d.StatusNo AND d.StatusSchema = 'Status'
                        INNER JOIN BAS.[User] e ON a.CreateBy = e.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingId", @" AND a.RoutingId = @RoutingId", RoutingId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeId", @" AND a.ModeId = @ModeId", ModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.CreateBy = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingType", @" AND a.RoutingType = @RoutingType", RoutingType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingName", @" AND a.RoutingName LIKE '%' + @RoutingName + '%'", RoutingName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingItemId", @" AND EXISTS (SELECT TOP 1 1 
                                                                                                                       FROM MES.RoutingItem x
                                                                                                                       WHERE x.RoutingId = a.RoutingId
                                                                                                                       AND x.RoutingItemId = @RoutingItemId)", RoutingItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND EXISTS (SELECT TOP 1 1 
                                                                                                                   FROM MES.RoutingItem x
                                                                                                                   INNER JOIN PDM.MtlItem xa ON x.MtlItemId = xa.MtlItemId
                                                                                                                   WHERE x.RoutingId = a.RoutingId
                                                                                                                   AND xa.MtlItemNo LIKE '%' + @MtlItemNo + '%')", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND EXISTS (SELECT TOP 1 1 
                                                                                                                   FROM MES.RoutingItem x
                                                                                                                   INNER JOIN PDM.MtlItem xa ON x.MtlItemId = xa.MtlItemId
                                                                                                                   WHERE x.RoutingId = a.RoutingId
                                                                                                                   AND xa.MtlItemName LIKE '%' + @MtlItemName + '%')", MtlItemName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (ProcessIds.Length > 0) {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessIds", @" AND a.RoutingId in(
                                                                                                                                SELECT a.RoutingId
                                                                                                                                FROM MES.Routing a
                                                                                                                                INNER JOIN MES.RoutingProcess b ON a.RoutingId = b.RoutingId
                                                                                                                                WHERE b.ProcessId IN @ProcessIds
                                                                                                                                GROUP BY a.RoutingId
                                                                                                                                HAVING COUNT(DISTINCT b.ProcessId) = @ProcessCount)", ProcessIds.Split(','));
                                                                                                                                dynamicParameters.Add("ProcessCount", ProcessIds.Split(',').Length);
                    }
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RoutingId DESC";
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

        #region //GetRoutingProcess -- 取得途程製程資料 -- Ann 2022-07-14
        public string GetRoutingProcess(int RoutingProcessId, int RoutingId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.RoutingProcessId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RoutingId, a.ProcessId, a.SortNumber, a.ProcessAlias, a.DisplayStatus
                        , a.NecessityStatus, a.ProcessCheckStatus, a.ProcessCheckType, a.PackageFlag, a.ConsumeFlag
                        , b.RoutingName, b.RoutingType, b.RoutingConfirm
                        , c.ProcessNo, c.ProcessName
                        , d.TypeName NecessityStatusName";
                    sqlQuery.mainTables =
                        @"FROM MES.RoutingProcess a
                        INNER JOIN MES.Routing b ON a.RoutingId = b.RoutingId
                        INNER JOIN MES.Process c ON a.ProcessId = c.ProcessId
                        INNER JOIN BAS.[Type] d ON a.NecessityStatus = d.TypeNo AND d.TypeSchema = 'RoutingProcess.NecessityStatus'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingProcessId", @" AND a.RoutingProcessId = @RoutingProcessId", RoutingProcessId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingId", @" AND a.RoutingId = @RoutingId", RoutingId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SortNumber";
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

        #region //GetRoutingItem -- 取得途程品號資料 -- Ann 2022-07-18
        public string GetRoutingItem(int RoutingItemId, int RoutingId, string MtlItemNo, string MtlItemName
            , int ControlId, string RoutingItemConfirm, int MtlItemId, int MoId, string RoutingName, int ModeId, string Edition, string RoutingConfirm, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.RoutingItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RoutingId, a.ControlId, a.MtlItemId, a.Status, a.RoutingItemConfirm
                        , b.MtlItemNo, b.MtlItemName, b.MtlItemSpec
                        , c.Edition, c.[Version], c.Cad2DFile, c.JmoFile, c.Cad2DFileAbsolutePath, c.JmoFileAbsolutePath
                        , (d.[FileName] + d.FileExtension) Cad2DFileInfo
                        , (e.[FileName] + e.FileExtension) JmoFileInfo
                        , f.RoutingName, f.ModeId
                        , g.ModeNo, g.ModeName";
                    if (MoId != -1) sqlQuery.columns += @", h.MoRoutingId, h.SortNumber";
                    sqlQuery.mainTables =
                        @"FROM MES.RoutingItem a
                        INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                        LEFT JOIN PDM.RdDesignControl c ON a.ControlId = c.ControlId
                        LEFT JOIN BAS.[File] d ON c.Cad2DFile = d.FileId
                        LEFT JOIN BAS.[File] e ON c.JmoFile = e.FileId
                        INNER JOIN MES.Routing f ON a.RoutingId = f.RoutingId
                        INNER JOIN MES.ProdMode g ON f.ModeId = g.ModeId";
                    if (MoId != -1) sqlQuery.mainTables += @" LEFT JOIN MES.MoRouting h ON a.RoutingItemId = h.RoutingItemId AND h.MoId  = @MoId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("MoId", MoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingItemId", @" AND a.RoutingItemId = @RoutingItemId", RoutingItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingId", @" AND a.RoutingId = @RoutingId", RoutingId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND b.MtlItemNo LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND b.MtlItemName LIKE '%' + @MtlItemName + '%'", MtlItemName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ControlId", @" AND a.ControlId = @ControlId", ControlId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingItemConfirm", @" AND a.RoutingItemConfirm = @RoutingItemConfirm", RoutingItemConfirm);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingName", @" AND f.RoutingName LIKE '%' + @RoutingName + '%'", RoutingName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeId", @" AND f.ModeId = @ModeId", ModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Edition", @" AND c.Edition LIKE '%' + @Edition + '%'", Edition);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingConfirm", @" AND f.RoutingConfirm = @RoutingConfirm", RoutingConfirm);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND (a.Status = @Status AND f.Status = @Status)", Status);
                    //if (MoId != -1) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MoId", @" AND h.MoId  = @MoId", MoId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RoutingItemId DESC";
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

        #region //GetRdDesignControl -- 取得研發設計圖版本控制資料 -- Ann 2024-09-10
        public string GetRdDesignControl(int ControlId, int DesignId, string Edition, string StartDate, string EndDate, int MtlItemId, string ReleasedStatus, string MtlItemNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ControlId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ControlId, a.Edition, a.[Version], a.Cause
                        , FORMAT(a.DesignDate, 'yyyy-MM-dd') DesignDate, a.ReleasedStatus
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd hh:mm:ss') CreateDate
                        , a.Cad3DFileAbsolutePath, a.Cad2DFileAbsolutePath, a.Pdf2DFileAbsolutePath, a.JmoFileAbsolutePath
                        , c.MtlItemId, c.CustomerMtlItemNo
                        , (
                          SELECT xa.FileId, (xa.[FileName] + xa.FileExtension) FileInfo
                          FROM BAS.[File] xa
                          WHERE xa.FileId = a.Cad3DFile
                          FOR JSON PATH, ROOT('data')
                        ) Cad3DFile
                        , (
                          SELECT xb.FileId, (xb.[FileName] + xb.FileExtension) FileInfo
                          FROM BAS.[File] xb
                          WHERE xb.FileId = a.Cad2DFile
                          FOR JSON PATH, ROOT('data')
                        ) Cad2DFile
                        , (
                          SELECT xc.FileId, (xc.[FileName] + xc.FileExtension) FileInfo
                          FROM BAS.[File] xc
                          WHERE xc.FileId = a.Pdf2DFile
                          FOR JSON PATH, ROOT('data')
                        ) Pdf2DFile
                        , (
                          SELECT xd.FileId, (xd.[FileName] + xd.FileExtension) FileInfo
                          FROM BAS.[File] xd
                          WHERE xd.FileId = a.JmoFile
                          FOR JSON PATH, ROOT('data')
                        ) JmoFile
                        , b.UserNo, b.UserName
                        , c.DesignId, c.MtlItemId";
                    sqlQuery.mainTables =
                        @"FROM PDM.RdDesignControl a
                        INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                        INNER JOIN PDM.RdDesign c ON a.DesignId = c.DesignId
                        INNER JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("DesignId", DesignId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ControlId", @" AND a.ControlId = @ControlId", ControlId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignId", @" AND a.DesignId = @DesignId", DesignId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Edition", @" AND a.Edition LIKE '%' + @Edition + '%'", Edition);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND c.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND d.MtlItemNo = @MtlItemNo", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ReleasedStatus", @" AND a.ReleasedStatus = @ReleasedStatus", ReleasedStatus);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.Edition, a.Version";
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

        #region //GetRoutingItemAttribute -- 取得途程品號加工屬性 -- Ann 2022-07-19
        public string GetRoutingItemAttribute(int RoutingItemId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.AttributeId
                            , a.RoutingItemId, a.AttributeItemId, a.AttributeValue, a.UpperLimit, a.LowerLimit
                            , b.ItemNo, b.ItemDesc
                            , c.ModeNo, c.ModeName
                            FROM MES.RoutingItemAttribute a
                            INNER JOIN PDM.RdAttributeItem b ON a.AttributeItemId = b.AttributeItemId
                            INNER JOIN MES.ProdMode c ON b.ModeId = c.ModeId
                            WHERE b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RoutingItemId", @" AND a.RoutingItemId = @RoutingItemId", RoutingItemId);

                    sql += @" ORDER BY a.AttributeId DESC";

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

        #region //GetRoutingItemProcess -- 取得途程品號流程卡資料 -- Ann 2022-07-25
        public string GetRoutingItemProcess(int RoutingItemId, string OrderBy)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ItemProcessId
                            , a.RoutingItemId, a.RoutingProcessId,a.AttrSetting
                            , ISNULL(a.RoutingItemProcessDesc, '') RoutingItemProcessDesc
                            , ISNULL(a.Remark, '') Remark
                            , a.DisplayStatus, a.CycleTime, a.MoveTime
                            , b.ProcessId, b.SortNumber, b.ProcessAlias
                            , c.ProcessNo, c.ProcessName
                            , d.RoutingName
                            , f.MtlItemNo
                            , a.ProcessingTime, a.WatingTime,b.ProcessCheckStatus
                            FROM MES.RoutingItemProcess a
                            INNER JOIN MES.RoutingProcess b ON a.RoutingProcessId = b.RoutingProcessId
                            INNER JOIN MES.Process c ON b.ProcessId = c.ProcessId
                            INNER JOIN MES.Routing d ON b.RoutingId = d.RoutingId
                            INNER JOIN MES.RoutingItem e ON a.RoutingItemId = e.RoutingItemId
                            INNER JOIN PDM.MtlItem f ON e.MtlItemId = f.MtlItemId
                            WHERE c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RoutingItemId", @" AND a.RoutingItemId = @RoutingItemId", RoutingItemId);

                    sql += @" ORDER BY b.SortNumber";

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

        #region //GetRoutingProcessItem -- 取得途程製程屬性檔 -- Ann 2022-08-12
        public string GetRoutingProcessItem(int RoutingProcessItemId, int RoutingProcessId, string ItemNo, string TypeName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.RoutingProcessItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RoutingProcessId, a.ItemNo, a.ChkUnique
                        , b.TypeNo, b.TypeName, b.TypeDesc";
                    sqlQuery.mainTables =
                        @"FROM MES.RoutingProcessItem a
                        INNER JOIN BAS.[Type] b ON a.ItemNo = b.TypeNo AND b.TypeSchema = 'RoutingProcessItem.ItemNo'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingProcessItemId", @" AND a.RoutingProcessItemId = @RoutingProcessItemId", RoutingProcessItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingProcessId", @" AND a.RoutingProcessId = @RoutingProcessId", RoutingProcessId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemNo", @" AND a.ItemNo LIKE '%' + @ItemNo + '%'", ItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeName", @" AND b.TypeName LIKE '%' + @TypeName + '%'", TypeName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RoutingProcessItemId DESC";
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

        #region //GetRipQcItem -- 取得途程製程量測項目 -- Ann 2022-09-05
        public string GetRipQcItem(int ItemProcessId, int QcItemId, string ItemNo, string ItemName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ItemProcessId, a.QcItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DesignValue, a.UpperValue, a.LimitValue
                        , b.ItemNo, b.ItemName, b.ItemDesc";
                    sqlQuery.mainTables =
                        @"FROM MES.RipQcItem a
                        INNER JOIN MES.QcItem b ON a.QcItemId = b.QcItemId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemProcessId", @" AND a.ItemProcessId = @ItemProcessId", ItemProcessId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcItemId", @" AND a.QcItemId = @QcItemId", QcItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemNo", @" AND b.ItemNo LIKE '%' + @ItemNo + '%'", ItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemName", @" AND b.ItemName LIKE '%' + @ItemName + '%'", ItemName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ItemProcessId DESC";
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

        #region //GetQcItem -- 取得製程量測項目 -- Ann 2022-09-05
        public string GetQcItem(int QcItemId, string ItemNo, string ItemName, int ProcessId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.ItemNo, a.ItemName, a.ItemDesc
                        , b.ProcessId";
                    sqlQuery.mainTables =
                        @"FROM MES.QcItem a
                        INNER JOIN MES.ProcessQcItem b ON a.QcItemId = b.QcItemId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessId", @" AND b.ProcessId = @ProcessId", ProcessId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcItemId DESC";
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

        #region //GetMoProcessChange -- 取得製令途程變更單 -- Shintokuro 2022.10.27
        public string GetMoProcessChange(int MpcId, int MoId, string StartDocDate, string EndDocDate, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MpcId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.MpcNo, a.MoId, a.MpcUserId,FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate, a.Remark, a.Status
                           ,(b1.WoErpPrefix + '-' + b1.WoErpNo + '(' + CONVERT ( Varchar , b.WoSeq ) + ')') MoNo
                          ,b2.MtlItemNo ,b2.MtlItemName
                          ,b3.LotStatus
                          ,c.UserName MpcUserName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.MoProcessChange a
                          INNER JOIN MES.ManufactureOrder b on a.MoId = b.MoId
                          INNER JOIN MES.WipOrder b1 on b.WoId = b1.WoId
                          INNER JOIN PDM.MtlItem b2 on b1.MtlItemId = b2.MtlItemId
                          INNER JOIN MES.MoSetting b3 on b.MoId = b3.MoId
                          INNER JOIN BAS.[User] c on a.MpcUserId = c.UserId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId=@CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MpcId", @" AND a.MpcId = @MpcId", MpcId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MoId", @" AND a.MoId = @MoId", MoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDocDate", @" AND a.DocDate >= @StartDocDate ", StartDocDate.Length > 0 ? Convert.ToDateTime(StartDocDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDocDate", @" AND a.DocDate <= @EndDocDate ", EndDocDate.Length > 0 ? Convert.ToDateTime(EndDocDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MpcId DESC";
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

        #region//GetMoProcess --取得製令製程 -- Shintokuro 2022.10.31
        public string GetMoProcess(int MoId, string ProcessNo, string ProcessName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MoProcessId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.ProcessAlias,a.MoId,a.SortNumber
                          ,b.ProcessId,b.ProcessNo,b.ProcessName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.MoProcess a
                          INNER JOIN MES.Process b on a.ProcessId = b.ProcessId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MoId", @" AND a.MoId = @MoId", MoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessNo", @" AND b.ProcessNo = @ProcessNo", ProcessNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessName", @" AND b.ProcessName = @ProcessName", ProcessName);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProcessAlias ASC";
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

        #region //GetMpcRoutingProcess -- 取得途程製程資料 -- Shintokuro 2022-10-31
        public string GetMpcRoutingProcess(int MpcRoutingProcessId, int MpcId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MpcRoutingProcessId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MpcId, a.MoProcessId, a.SortNumber, a.ProcessAlias, a.DisplayStatus
                        , a.NecessityStatus, a.RandomStatus
                        , b.Status
                        , d.TypeName NecessityStatusName
                        , e.TypeName RandomStatusName";
                    sqlQuery.mainTables =
                        @"FROM MES.MpcRoutingProcess a
                        INNER JOIN MES.MoProcessChange b ON a.MpcId = b.MpcId
                        INNER JOIN MES.MoProcess c ON a.MoProcessId = c.MoProcessId
                        INNER JOIN BAS.[Type] d ON a.NecessityStatus = d.TypeNo AND d.TypeSchema = 'RoutingProcess.NecessityStatus'
                        INNER JOIN BAS.[Type] e ON a.RandomStatus = e.TypeNo AND e.TypeSchema = 'RoutingProcess.RandomStatus'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MpcRoutingProcessId", @" AND a.MpcRoutingProcessId = @MpcRoutingProcessId", MpcRoutingProcessId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MpcId", @" AND a.MpcId = @MpcId", MpcId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SortNumber";
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

        #region //GetMpcBarcode -- 取得途程變更單綁定條碼 -- Shintokuro 2022-11-01
        public string GetMpcBarcode(int MpcId, int MpcBarcodeId, string BarcodeNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MpcBarcodeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MpcId, a.BarcodeNo";
                    sqlQuery.mainTables =
                        @"FROM MES.MpcBarcode a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MpcId", @" AND a.MpcId = @MpcId", MpcId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BarcodeNo", @" AND a.BarcodeNo = @BarcodeNo", BarcodeNo);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MpcBarcodeId";
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

        #region //GetMpcBarcodeRouting -- 取得途程製程資料 -- Shintokru 2022-11-04
        public string GetMpcBarcodeRouting(int MpcBarcodeId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MpcBarcodeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RoutingId, a.ProcessId, a.SortNumber, a.ProcessAlias, a.DisplayStatus
                        , a.NecessityStatus, a.RandomStatus
                        , b.RoutingName, b.RoutingType, b.RoutingConfirm
                        , c.ProcessNo, c.ProcessName
                        , d.TypeName NecessityStatusName
                        , e.TypeName RandomStatusName";
                    sqlQuery.mainTables =
                        @"FROM MES.RoutingProcess a
                        INNER JOIN MES.Routing b ON a.RoutingId = b.RoutingId
                        INNER JOIN MES.Process c ON a.ProcessId = c.ProcessId
                        INNER JOIN BAS.[Type] d ON a.NecessityStatus = d.TypeNo AND d.TypeSchema = 'RoutingProcess.NecessityStatus'
                        INNER JOIN BAS.[Type] e ON a.RandomStatus = e.TypeNo AND e.TypeSchema = 'RoutingProcess.RandomStatus'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MpcBarcodeId", @" AND a.MpcBarcodeId = @MpcBarcodeId", MpcBarcodeId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SortNumber";
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

        #region //GetModeProcess -- 取得生產模式製程資料 -- Ann 2022-12-01
        public string GetModeProcess(int ProcessId, string ProcessNo, string ProcessName, string Status, int ProdMode
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ProcessId, a.ProcessNo, a.ProcessName, a.ProcessDesc, a.Status
                            , (a.ProcessNo + '-' + a.ProcessName) ProcessWithText
                            , c.ParameterNo, b.ParameterId
                            FROM MES.Process a
                            INNER JOIN MES.ProcessParameter b ON a.ProcessId = b.ProcessId
                            LEFT JOIN BAS.Parameter c ON a.ProcessNo = c.ParameterValue
                            WHERE a.CompanyId = @CompanyId
                            AND b.Status = 'A'";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ModeId", @" AND b.ModeId = @ModeId", ProdMode);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProcessNo", @" AND a.ProcessNo LIKE '%' + @ProcessNo + '%'", ProcessNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProcessName", @" AND a.ProcessName LIKE '%' + @ProcessName + '%'", ProcessName);

                    sql += @" ORDER BY a.ProcessNo";
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

        #region //GetRoutingProcessItemList -- 取得途程製程資料屬性 -- Ding 2023-04-16
        public string GetRoutingProcessItemList(int RoutingProcessId, int RoutingId, string ProcessIdList
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.RoutingProcessId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RoutingId, a.ProcessId, a.SortNumber, a.ProcessAlias, a.DisplayStatus
                        , a.NecessityStatus, a.ProcessCheckStatus, a.ProcessCheckType
                        , b.RoutingName, b.RoutingType, b.RoutingConfirm
                        , c.ProcessNo, c.ProcessName
                        , d.TypeName NecessityStatusName";
                    sqlQuery.mainTables =
                        @"FROM MES.RoutingProcess a
                        INNER JOIN MES.Routing b ON a.RoutingId = b.RoutingId
                        INNER JOIN MES.Process c ON a.ProcessId = c.ProcessId
                        INNER JOIN BAS.[Type] d ON a.NecessityStatus = d.TypeNo AND d.TypeSchema = 'RoutingProcess.NecessityStatus'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingProcessId", @" AND a.RoutingProcessId = @RoutingProcessId", RoutingProcessId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoutingId", @" AND a.RoutingId = @RoutingId", RoutingId);
                    if (!ProcessIdList.Equals("")) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessIdList", @" AND a.ProcessId IN ( @ProcessIdList)", ProcessIdList);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SortNumber";
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

        #region//GetPreFlowCardPdf 途程流程卡預覽
        public string GetPreFlowCardPdf(int RoutingItemId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT c.SortNumber, c.ProcessAlias, b.RoutingItemProcessDesc, b.Remark,d.MtlItemNo,d.MtlItemName,d.MtlItemSpec,f.CustomerMtlItemNo,e.Edition
                            FROM MES.RoutingItem a
                                LEFT JOIN MES.RoutingItemProcess b ON a.RoutingItemId=b.RoutingItemId
                                LEFT JOIN MES.RoutingProcess c ON b.RoutingProcessId=c.RoutingProcessId
                                LEFT JOIN PDM.MtlItem d ON d.MtlItemId=a.MtlItemId
                                LEFT JOIN PDM.RdDesignControl e ON e.ControlId=a.ControlId
                                LEFT JOIN PDM.RdDesign f ON f.DesignId=e.DesignId
                             WHERE b.DisplayStatus='Y' 
                             AND a.RoutingItemId=@RoutingItemId
                            ORDER BY c.SortNumber ASC";
                    dynamicParameters.Add("@RoutingItemId", RoutingItemId);
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

        #region//ApiAiAutoRoutingV2 呼叫 人工智慧研發部 API 自動新建 一組全套或半套途程。按秉丞指導:不使用 async, 業務邏輯全部改到DA,  by MarkChen 2023.08.24
        public string ApiAiAutoRoutingV2(int MtlItemId, int ControlId)
        {
            Console.WriteLine("RoutingDA => ApiAiAutoRoutingV2(int MtlItemId, int ControlId)");
            #region //[開始] NOTE by MarkChen, 使用 工廠模式 (Factory Pattern), 將複雜部份推到 DA
            var flow = new RoutingDA.ResultFactory().Create(MtlItemId, ControlId);
            #endregion

            #region //Step1 確認 研發設計圖版本控制 基礎資料, 是否適合調用 人工智慧研發部 API
            //CASE_A_01 : 查無 RdDesginControl
            if (flow.Step1 == null)
            {
                flow.Status = RoutingDA.Cases.CASE_A_01_RdDesignControl_NOT_FOUND;
                return JsonConvert.SerializeObject(flow); ;
            }

            //"CASE_A_02 : 無 FileId",
            if (flow.Step1.Cad2DFile == null)
            {
                flow.Status = RoutingDA.Cases.CASE_A_02_NO_FILEID;
                return JsonConvert.SerializeObject(flow); ;
            }
            flow.Input.FileId = (int)flow.Step1.Cad2DFile;
            #endregion

            #region //Step2 調用 人工智慧研發部 API, 待命調整為內部調用 或 以實際路徑調用

            using (HttpClient httpClient = new HttpClient())
            {
                var loginUri = "http://192.168.20.97/WorkManagement/Home/tx_login_check";

                var loginContent = new MultipartFormDataContent();
                loginContent.Add(new StringContent("{\"job_no\":\"mesapi\",\"password\":\"JmoAdmin123456\",\"remember_me\":true}"), "DATA");

                var loginResponse = httpClient.PostAsync(loginUri, loginContent).Result;
                if (!loginResponse.IsSuccessStatusCode)
                {
                    flow.Status = RoutingDA.Cases.CASE_B_01_NOT_ABLE_TO_VISIT_AI_API;
                    return JsonConvert.SerializeObject(flow);
                }

                string content = loginResponse.Content.ReadAsStringAsync().Result;
                JObject jsonResponse = JObject.Parse(content);

                string status = (string)jsonResponse["status"];
                string msg = (string)jsonResponse["msg"];
                string retdata = (string)jsonResponse["retdata"];

                if (status != "success")
                {
                    flow.Status = RoutingDA.Cases.CASE_B_02_NOT_ABLE_TO_LOGIN_AI_API;
                    flow.Debug = msg;
                    return JsonConvert.SerializeObject(flow);
                }

                var requestUri = "http://192.168.20.97/WorkManagement/Home/api_retrieve_file";

                var formContent = new MultipartFormDataContent();
                formContent.Add(new StringContent(flow.Step1.Cad2DFile.ToString()), "file_id");

                var response = httpClient.PostAsync(requestUri, formContent).Result;
                if (!response.IsSuccessStatusCode)
                {
                    flow.Status = RoutingDA.Cases.CASE_B_03_FAILED_TO_GET_PARSE;
                    return JsonConvert.SerializeObject(flow);
                }

                flow.Step2 = JsonConvert.DeserializeObject<RoutingDA.Root>(response.Content.ReadAsStringAsync().Result);

                if (flow.Step2.Python_Ret.Status != "success")
                {
                    flow.Status = RoutingDA.Cases.CASE_B_04_PARSED_RESULT_NOT_USEFUL;
                    flow.Debug = "有解析, 但不能為本案所用, 目前只限 mode = 1,  JMO-A-001 模仁加工";
                    return JsonConvert.SerializeObject(flow);
                }
            }



            #endregion

            #region //Step3 判斷新增模式是全套途程或半套途程
            // NOTE by Mark, 合併簡潔寫法, 此段仍可優化
            flow.Step3.RoutingName = GetRoutingNameV3(flow.Step2, (int)flow.Step1.Cad2DFile);
            flow.Step3.SameNameRouting = GetRoutingByName(flow.Step3.RoutingName);

            //判斷是否已有 同名製程, 
            //  發生情況, 例如 /ApiAiAutoRouting?MtlItemId=264065&ControlId=3987 , Call了兩次
            if (flow.Step3.SameNameRouting.RoutingId > 0)
            {
                flow.Status = RoutingDA.Cases.CASE_C_01_EXISTING_ROUTINGNAME; //已有同名製程
                flow.Step3.IsAnySameRoutingName = true;
                return JsonConvert.SerializeObject(flow); ;
            }

            //在JSON, 可以看起來 乾淨! 不會展開沒有值的 { }
            flow.Step3.SameNameRouting = null;
            flow.Step3.IsAnySameRoutingName = false;

            // DOING... 使用密碼學比較有沒有相同途程
            RoutingDA.IsExistingRoutingResponse r1 = IsExistingRoutingAdvAsync(flow.Step2, (int)flow.Step1.Cad2DFile);
            flow.Step3.EntireCompanyRoutingCnt = r1.EntireCompanyRoutingCnt;

            flow.Step3.ProcessCnt = r1.ProcessCnt;
            flow.Step3.SameProcessCntRoutingCnt = r1.SameProcessCntRoutingCnt;
            flow.Step3.strGoingToAdd = r1.strGoingToAdd;
            flow.Step3.IsAnySameProcessList = r1.IsAnySameProcessList;
            flow.Step3.SameProcessListDifferentRoutingName = r1.SameProcessListDifferentRoutingName;
            #endregion

            #region //Step4 執行 新增 全套途程 或 半套途程
            if (!flow.Step3.IsAnySameProcessList)
            {
                //新增  上邊 右邊  左邊 
                flow.Status = RoutingDA.Cases.CASE_D_01_TO_CREATE_ENTIRE_ROUTING;
                flow.Debug = "新增  上邊 右邊  左邊 ";

                flow.Step4.DebugMsg = "新增  上邊 右邊  左邊, calling  ExecuteTransaction_CASE_D_01... ";
                flow = ExecuteTransaction_CASE_D_01(flow, MtlItemId, ControlId);
            }
            else
            {
                //只新增 左邊 
                flow.Status = RoutingDA.Cases.CASE_D_02_TO_CREATE_ROUTING_ITEM_ONLY;
                flow.Debug = "只新增 左邊 .";

                flow.Step4.DebugMsg = "新增    左邊, calling  ExecuteTransaction_CASE_D_02... ";
                flow = ExecuteTransaction_CASE_D_02(flow, MtlItemId, ControlId);
            }
            #endregion

            #region //[結束] 回傳結果, 可以 不顯示DEBUG 細節
            //Clean Up Debug info
            flow.Cases = null;
            flow.Step1 = null;
            flow.Step2 = null;
            flow.Step3 = null;
            flow.Step4 = null;

            //flow.Step4.RoutingProcessList = null;
            return JsonConvert.SerializeObject(flow); ;
            #endregion

        }
        #endregion

        #region//ApiAiAutoRoutingV2 呼叫 人工智慧研發部 API 自動新建 一組全套或半套途程。按經理 08/25 指導: 讓用戶可以[途程製程維護],  by MarkChen 2023.08.25
        public string ApiAiAutoRoutingV2Step1(int MtlItemId, int ControlId)
        {
            //
            // 
            //
            Console.WriteLine("RoutingDA => ApiAiAutoRoutingV2(int MtlItemId, int ControlId)");
            #region //[開始] NOTE by MarkChen, 使用 工廠模式 (Factory Pattern), 將複雜部份推到 DA
            var flow = new RoutingDA.ResultFactory().Create(MtlItemId, ControlId);
            #endregion

            #region //Step1 確認 研發設計圖版本控制 基礎資料, 是否適合調用 人工智慧研發部 API
            //CASE_A_01 : 查無 RdDesginControl
            if (flow.Step1 == null)
            {
                flow.Status = RoutingDA.Cases.CASE_A_01_RdDesignControl_NOT_FOUND;
                return JsonConvert.SerializeObject(flow); ;
            }

            //"CASE_A_02 : 無 FileId",
            if (flow.Step1.Cad2DFile == null)
            {
                flow.Status = RoutingDA.Cases.CASE_A_02_NO_FILEID;
                return JsonConvert.SerializeObject(flow); ;
            }
            flow.Input.FileId = (int)flow.Step1.Cad2DFile;
            #endregion

            #region //Step2 調用 人工智慧研發部 API, 待命調整為內部調用 或 以實際路徑調用

            using (HttpClient httpClient = new HttpClient())
            {
                var loginUri = "http://192.168.20.97/WorkManagement/Home/tx_login_check";

                var loginContent = new MultipartFormDataContent();
                loginContent.Add(new StringContent("{\"job_no\":\"mesapi\",\"password\":\"JmoAdmin123456\",\"remember_me\":true}"), "DATA");

                var loginResponse = httpClient.PostAsync(loginUri, loginContent).Result;
                if (!loginResponse.IsSuccessStatusCode)
                {
                    flow.Status = RoutingDA.Cases.CASE_B_01_NOT_ABLE_TO_VISIT_AI_API;
                    return JsonConvert.SerializeObject(flow);
                }

                string content = loginResponse.Content.ReadAsStringAsync().Result;
                JObject jsonResponse = JObject.Parse(content);

                string status = (string)jsonResponse["status"];
                string msg = (string)jsonResponse["msg"];
                string retdata = (string)jsonResponse["retdata"];

                if (status != "success")
                {
                    flow.Status = RoutingDA.Cases.CASE_B_02_NOT_ABLE_TO_LOGIN_AI_API;
                    flow.Debug = msg;
                    return JsonConvert.SerializeObject(flow);
                }

                var requestUri = "http://192.168.20.97/WorkManagement/Home/api_retrieve_file";

                var formContent = new MultipartFormDataContent();
                formContent.Add(new StringContent(flow.Step1.Cad2DFile.ToString()), "file_id");

                var response = httpClient.PostAsync(requestUri, formContent).Result;
                if (!response.IsSuccessStatusCode)
                {
                    flow.Status = RoutingDA.Cases.CASE_B_03_FAILED_TO_GET_PARSE;
                    return JsonConvert.SerializeObject(flow);
                }

                flow.Step2 = JsonConvert.DeserializeObject<RoutingDA.Root>(response.Content.ReadAsStringAsync().Result);

                if (flow.Step2.Python_Ret.Status != "success")
                {
                    flow.Status = RoutingDA.Cases.CASE_B_04_PARSED_RESULT_NOT_USEFUL;
                    flow.Debug = "有解析, 但不能為本案所用, 目前只限 mode = 1,  JMO-A-001 模仁加工";
                    return JsonConvert.SerializeObject(flow);
                }
            }



            #endregion

            #region //Step3 判斷新增模式是全套途程或半套途程
            // NOTE by Mark, 合併簡潔寫法, 此段仍可優化
            flow.Step3.RoutingName = GetRoutingNameV3(flow.Step2, (int)flow.Step1.Cad2DFile);
            flow.Step3.SameNameRouting = GetRoutingByName(flow.Step3.RoutingName);

            //判斷是否已有 同名製程, 
            //  發生情況, 例如 /ApiAiAutoRouting?MtlItemId=264065&ControlId=3987 , Call了兩次
            if (flow.Step3.SameNameRouting.RoutingId > 0)
            {
                flow.Status = RoutingDA.Cases.CASE_C_01_EXISTING_ROUTINGNAME; //已有同名製程
                flow.Step3.IsAnySameRoutingName = true;
                return JsonConvert.SerializeObject(flow); ;
            }

            //在JSON, 可以看起來 乾淨! 不會展開沒有值的 { }
            flow.Step3.SameNameRouting = null;
            flow.Step3.IsAnySameRoutingName = false;

            // DOING... 使用密碼學比較有沒有相同途程
            RoutingDA.IsExistingRoutingResponse r1 = IsExistingRoutingAdvAsync(flow.Step2, (int)flow.Step1.Cad2DFile);
            flow.Step3.EntireCompanyRoutingCnt = r1.EntireCompanyRoutingCnt;

            flow.Step3.ProcessCnt = r1.ProcessCnt;
            flow.Step3.SameProcessCntRoutingCnt = r1.SameProcessCntRoutingCnt;
            flow.Step3.strGoingToAdd = r1.strGoingToAdd;
            flow.Step3.IsAnySameProcessList = r1.IsAnySameProcessList;
            flow.Step3.SameProcessListDifferentRoutingName = r1.SameProcessListDifferentRoutingName;
            #endregion

            #region //Step4 執行 新增 全套途程 或 半套途程
            flow.Status = RoutingDA.Cases.CASE_D_00;
            flow.Debug = "要能顯示 python 回傳的...";

            #endregion

            #region //[結束] 回傳結果, 可以 不顯示DEBUG 細節
            //Clean Up Debug info
            //flow.Cases = null;
            //flow.Step1 = null;F
            //flow.Step2 = null;
            //flow.Step3 = null;
            //flow.Step4 = null;

            //flow.Step4.RoutingProcessList = null;
            return JsonConvert.SerializeObject(flow); ;
            #endregion

        }
        #endregion

        #region//ApiAiAutoRoutingV2Step2 自動新建 一組全套途程(不含右2左2), by MarkChen 2023.09.01
        public string ApiAiAutoRoutingV2Step2(int MtlItemId, int ControlId, string ProcessJsonString)
        {
            #region //[開始] 使用 工廠模式 (Factory Pattern), 將複雜部份推到DA
            var flow = new RoutingDA.ResultFactory().CreateV2(MtlItemId, ControlId, ProcessJsonString);
            #endregion

            #region //Step1 確認 研發設計圖版本控制 基礎資料, 是否適合調用 人工智慧研發部 API
            //CASE_A_01: 查無 RdDesginControl
            if (flow.Step1 == null)
            {
                flow.Status = RoutingDA.Cases.CASE_A_01_RdDesignControl_NOT_FOUND;
                return JsonConvert.SerializeObject(flow); ;
            }

            //"CASE_A_02 : 無 FileId",
            if (flow.Step1.Cad2DFile == null)
            {
                flow.Status = RoutingDA.Cases.CASE_A_02_NO_FILEID;
                return JsonConvert.SerializeObject(flow); ;
            }
            flow.Input.FileId = (int)flow.Step1.Cad2DFile;
            #endregion



            //設置頂級變量
            flow.X_MtlItemId = MtlItemId;
            flow.Y_ControlId = ControlId;
            flow.Z_FilelId = flow.Input.FileId;

            //便於在流覽器DEBUG查看
            flow.StrAiProcess = ProcessJsonString;
            try
            {
                //將 人工智慧研發部 API 解析結果放到指定變量, 同上,便於在流覽器DEBUG查看 
                AiProcessWrapper wrapper = JsonConvert.DeserializeObject<AiProcessWrapper>(ProcessJsonString);
                List<AiProcess> aiProcesses = wrapper.ProcessInfo;
                flow.listAiProcess = aiProcesses;
            }
            catch (Exception ex)
            {
                flow.Debug = ex.Message;
            }

            //先確定命名規則的途程名稱
            flow.RoutingName = GetRoutingNameV4(flow.listAiProcess, (int)flow.Step1.Cad2DFile);

            //以 SQL Transaction 新建整套途程, 但目前不含右1及左1, 因無UI給用戶增刪改
            flow = ExecuteTransaction_Meeting0901(flow, MtlItemId, ControlId);

            #region //[結束] 回傳結果, 可以 不顯示DEBUG 細節
            //flow.Cases = null;
            //flow.Step1 = null;F
            //flow.Step2 = null;
            //flow.Step3 = null;
            //flow.Step4 = null;
            //flow.Step4.RoutingProcessList = null;
            return JsonConvert.SerializeObject(flow); ;
            #endregion

        }
        #endregion

        #region //Type for 自動流程卡專案 ApiAiAutoRouting -- MarkChen 2023-08-22
        public interface IResultFactory
        {
            RdDesignControlResult Create(int MtlItemId, int ControlId);
            RdDesignControlResult CreateV2(int MtlItemId, int ControlId, string ProcessJsonString);
        }
        public class ResultFactory : IResultFactory
        {
            public RdDesignControlResult Create(int MtlItemId, int ControlId)
            {
                var r = new RoutingDA();
                var flow = new RoutingDA.RdDesignControlResult();
                // NOTE by Mark, 2023-08-22
                // 將裡層的 object 也建出來, 就不必在 biz logic 處理
                flow.Cases = RoutingDA.Cases.GetAll();

                var rdDesignControlInput = new RoutingDA.RdDesignControlInput();

                rdDesignControlInput.MtlItemId = MtlItemId;
                rdDesignControlInput.ControlId = ControlId;
                flow.Input = rdDesignControlInput;

                RdDesignControl rdDesignControl = r.GetRdDesignControl(MtlItemId, ControlId);
                flow.Step1 = rdDesignControl;

                RoutingDA.Root rootObject2 = new RoutingDA.Root();
                flow.Step2 = rootObject2;

                RoutingDA.IsExistingRoutingResponse caseC = new RoutingDA.IsExistingRoutingResponse();
                flow.Step3 = caseC;
                RoutingDA.CaseDResponse caseDResponse = new RoutingDA.CaseDResponse();
                flow.Step4 = caseDResponse;
                return flow;
            }

            public RdDesignControlResult CreateV2(int MtlItemId, int ControlId, string ProcessJsonString)
            {
                var r = new RoutingDA();
                var flow = new RoutingDA.RdDesignControlResult();
                // NOTE by Mark, 2023-08-22
                // 將裡層的 object 也建出來, 就不必在 biz logic 處理
                flow.Cases = RoutingDA.Cases.GetAll();

                var rdDesignControlInput = new RoutingDA.RdDesignControlInput();

                rdDesignControlInput.MtlItemId = MtlItemId;
                rdDesignControlInput.ControlId = ControlId;
                flow.Input = rdDesignControlInput;

                RdDesignControl rdDesignControl = r.GetRdDesignControl(MtlItemId, ControlId);
                flow.Step1 = rdDesignControl;

                RoutingDA.Root rootObject2 = new RoutingDA.Root();
                flow.Step2 = rootObject2;

                RoutingDA.IsExistingRoutingResponse caseC = new RoutingDA.IsExistingRoutingResponse();
                flow.Step3 = caseC;
                RoutingDA.CaseDResponse caseDResponse = new RoutingDA.CaseDResponse();
                flow.Step4 = caseDResponse;
                return flow;
            }
        }
        public class DboVCompanyRoutingProcessCnt
        {

            /*
                CREATE VIEW V_CompanyRoutingProcessCnt as 
                SELECT CompanyId,T2.RoutingId,Count(*) ProcessCnt 
                FROM 
                MES.RoutingProcess  T1
                , MES.Routing  T2
                WHERE T1.RoutingId=T2.RoutingId
                AND T2.CompanyId=2
                GROUP BY CompanyId, T2.RoutingId
             */



            public int CompanyId { get; set; }
            public int RoutingId { get; set; }
            public int ProcessCnt { get; set; }

        }
        public string GetRoutingNameV3(Root rootObject, int fileId)
        {
            string result = $"【AI-{fileId}】";
            foreach (var x in rootObject.Python_Ret.Data.RountingProcess.OrderBy(a => a.SortNumber))
            {
                //string str = "(" + x.SortNumber + ")" + x.ProcessAlias;
                string str = x.ProcessAlias + "|";
                result += str;


            }
            var result2 = result.Substring(0, result.Length - 1);
            return result2;
        }
        public string GetRoutingNameV4(List<AiProcess> aiProcesses, int fileId)
        {
            string result = $"【AI-{fileId}】";
            foreach (var x in aiProcesses.OrderBy(a => a.SortNumber))
            {
                //string str = "(" + x.SortNumber + ")" + x.ProcessAlias;
                string str = x.ProcessAlias + "|";
                result += str;


            }
            var result2 = result.Substring(0, result.Length - 1);
            return result2;
        }
        public class RdDesignControl
        {
            public int MtlItemId { get; set; }
            public int ControlId { get; set; }
            public string MtlItemNo { get; set; }
            public string MtlItemName { get; set; }
            public int DesignId { get; set; }
            public int? Cad2DFile { get; set; }
            // ... other properties
        }

        public class RdDesignControlResult
        {
            public List<AiProcess> listAiProcess { get; set; } //NOTE by Mark, 09/01
            public List<string> Cases { get; set; }
            public string Status { get; set; }
            public string Debug { get; set; }
            //int MtlItemId, int ControlId
            public int X_MtlItemId { get; set; }
            public int Y_ControlId { get; set; }
            public int Z_FilelId { get; set; }//RoutingName
            public string RoutingName { get; set; }
            public string StrAiProcess { get; set; }
            public RdDesignControlInput Input { get; set; }
            public RdDesignControl Step1 { get; set; }
            public Root Step2 { get; set; }
            public IsExistingRoutingResponse Step3 { get; set; }
            public CaseDResponse Step4 { get; set; }

        }
        public class CaseDResponse
        {
            public string CaseType { get; set; }
            public string DebugMsg { get; set; }
            public string DebugSQL1 { get; set; }
            public string DebugSQL2 { get; set; }
            public string DebugSQL3 { get; set; }
            public string DebugSQL4 { get; set; }
            public string DebugSQL5 { get; set; }
            public string ConnStr { get; set; }
            public DateTime TransactionStart { get; set; }
            public DateTime TransactionEnd { get; set; }
            public string TransactionResult { get; set; }

            public string RoutingName { get; set; }
            public MesRouting Level1Routing { get; set; }
            public int Level1RoutingId { get; set; }
            public int Level2RoutingProcessId { get; set; }
            public string Level2RoutingProcessNo { get; set; }

            public int Level3RoutingProcessItemId { get; set; }
            public int Level4RoutingItemId { get; set; }
            public int Level5RoutingItemProcessId { get; set; }
            public int ProcessId { get; set; }
            public string ProcessNo { get; set; }
            public List<MesProcess> ProcessList { get; set; }// 全套時使用
            public List<MesRoutingProcess> RoutingProcessList { get; set; }// 半套時使用
            public int RoutingId { get; set; }
            public int RoutingItemId { get; set; }
            public int TemplateRoutingItemId { get; set; }
            public List<MesRoutingItemProcess> TemplateRoutingItemProcess { get; set; }// 半套時使用
            //public int RoutingId { get; set; }

        }

        public class BasFile
        {
            //[Key]
            //[DatabaseGenerated(DatabaseGeneratedOption.Identity)]
            public int FileId { get; set; }

            //[Required]
            public int CompanyId { get; set; }

            //[Required]
            public string FileName { get; set; }

            //[Required]
            public byte[] FileContent { get; set; }

            //[Required]
            public string FileExtension { get; set; }

            //[Required]
            public int FileSize { get; set; }

            //[Required]
            public string ClientIP { get; set; }

            //[Required]
            public string Source { get; set; }

            //[Required]
            public string DeleteStatus { get; set; }

            public DateTime CreateDate { get; set; }

            public DateTime? LastModifiedDate { get; set; }

            //[Required]
            public int CreateBy { get; set; }

            public int? LastModifiedBy { get; set; }

            //public ICollection<BasCompany> BasCompanys { get; set; }

            //public BasCompany BasCompany { get; set; }

        }
        public class MesRoutingItemProcess
        {
            //[Key]
            //[DatabaseGenerated(DatabaseGeneratedOption.Identity)]
            public int ItemProcessId { get; set; }
            public int RoutingItemId { get; set; }
            public int RoutingProcessId { get; set; }

            //這版本不接受字串為NULL
            //public string? RoutingItemProcessDesc { get; set; }
            public string RoutingItemProcessDesc { get; set; }

            public int CycleTime { get; set; } = 0; // Default value
            public int MoveTime { get; set; } = 0; // Default value
                                                   //這版本不接受字串為NULL
                                                   //public string? Remark { get; set; }
            public string Remark { get; set; }
            public string DisplayStatus { get; set; }
            //      public string? AttrSetting { get; set; }
            public DateTime CreateDate { get; set; } = DateTime.Now; // Default value
            public DateTime? LastModifiedDate { get; set; }
            public int CreateBy { get; set; }
            public int? LastModifiedBy { get; set; }

            // Navigation properties (if needed) can be added here
            // e.g., public virtual RoutingItem RoutingItem { get; set; }
            //       public virtual RoutingProcess RoutingProcess { get; set; }
        }
        public class MesProcess
        {
            //[Key]
            //[DatabaseGenerated(DatabaseGeneratedOption.Identity)]
            public int ProcessId { get; set; }

            //[Required]
            public int CompanyId { get; set; }

            public int? DepartmentId { get; set; }

            //[Required]
            //[MaxLength(50)]
            //  [Index("AK1_Process", IsUnique = true, Order = 2)]
            public string ProcessNo { get; set; }

            //[Required]
            //[MaxLength(100)]
            //   [Index("AK2_Process", IsUnique = true, Order = 2)]
            public string ProcessName { get; set; }

            //[Required]
            //[MaxLength(100)]
            public string ProcessDesc { get; set; }

            //[Required]
            //[MaxLength(2)]
            public string Status { get; set; }

            //[Required]
            //[Column(TypeName = "datetime")]
            //[DatabaseGenerated(DatabaseGeneratedOption.Computed)]
            public DateTime CreateDate { get; set; }

            //[Column(TypeName = "datetime")]
            public DateTime? LastModifiedDate { get; set; }

            //[Required]
            public int CreateBy { get; set; }

            public int? LastModifiedBy { get; set; }

            // Navigation Properties
            //[ForeignKey("CompanyId")]
            //public virtual Company Company { get; set; }  // Assuming you have a Company model
        }

        public class MesRouting
        {
            public int RoutingId { get; set; }
            public int CompanyId { get; set; }
            public int ModeId { get; set; }
            public string RoutingType { get; set; }

            public string RoutingName { get; set; }

            public string Status { get; set; }

            public string RoutingConfirm { get; set; }

            public DateTime CreateDate { get; set; }

            public DateTime? LastModifiedDate { get; set; }

            public int CreateBy { get; set; }

            public int? LastModifiedBy { get; set; }
        }

        public class MesRoutingProcess
        {
            public int RoutingProcessId { get; set; }

            public int RoutingId { get; set; }

            public int ProcessId { get; set; }

            public int SortNumber { get; set; }

            public string ProcessAlias { get; set; }

            public string DisplayStatus { get; set; }

            public string NecessityStatus { get; set; }

            public string ProcessCheckStatus { get; set; }

            public string ProcessCheckType { get; set; }

            public string PackageFlag { get; set; }

            public string Status { get; set; }

            public DateTime CreateDate { get; set; }

            public DateTime? LastModifiedDate { get; set; }

            public int CreateBy { get; set; }

            public int? LastModifiedBy { get; set; }

            //public virtual Routing Routing { get; set; }
            //public virtual Process Process { get; set; }

            /*
                public string ProcessNo { get; set; } //製程編號
                public string ProcessAlias { get; set; }//製程別名
                public int SortNumber { get; set; }//製程順序
                public string ProcessCheckStatus { get; set; } //支援工程檢
                public string ProcessCheckType { get; set; }//工程檢頻率
                public string DisplayStatus { get; set; } //是否顯示在流程卡
                public string NecessityStatus { get; set; }//必要過站
             */
            public override int GetHashCode()
            {
                int hash = 17; // Starting value for the hash code

                // Include all properties you want to consider for the hash, except RoutingId and ProcessId
                //  hash = hash * 23 + ProcessNo.GetHashCode();
                hash = hash * 23 + SortNumber.GetHashCode();
                hash = hash * 23 + ProcessAlias.GetHashCode();
                hash = hash * 23 + DisplayStatus.GetHashCode();
                hash = hash * 23 + NecessityStatus.GetHashCode();
                hash = hash * 23 + ProcessCheckStatus.GetHashCode();
                hash = hash * 23 + (ProcessCheckType?.GetHashCode() ?? 0);
                //hash = hash * 23 + PackageFlag.GetHashCode();
                //hash = hash * 23 + Status.GetHashCode();
                //hash = hash * 23 + CreateDate.GetHashCode();
                //hash = hash * 23 + (LastModifiedDate?.GetHashCode() ?? 0);
                //hash = hash * 23 + CreateBy.GetHashCode();
                //hash = hash * 23 + (LastModifiedBy?.GetHashCode() ?? 0);

                return hash;
            }
        }

        public class MesRoutingProcessItem
        {
            public int RoutingProcessItemId { get; set; }
            public int RoutingProcessId { get; set; }
            public string ItemNo { get; set; }
            public char ChkUnique { get; set; }
            public DateTime CreateDate { get; set; }
            public DateTime? LastModifiedDate { get; set; }
            public int CreateBy { get; set; }
            public int? LastModifiedBy { get; set; }

            // Navigation property for the related RoutingProcess
            //    public virtual RoutingProcess RoutingProcess { get; set; }
            //}
        }


        public class SubCallResponse
        {
            public bool IsExisting { get; set; }
            public string Status { get; set; }
            public string Msg { get; set; }
            public string RoutingName { get; set; }
            public MesRouting ExistingMesRouting { get; set; }
        }

        public class Cases
        {
            /// <summary>
            /// CASE_A_00 : === 研發設計圖版本控制 ===
            /// </summary>
            public static string CASE_A_00 = "CASE_A_00 : === 研發設計圖版本控制 ===";
            /// <summary>
            /// CASE_A_01 : 查無 RdDesginControl
            /// </summary>
            public static string CASE_A_01_RdDesignControl_NOT_FOUND = "CASE_A_01 : 查無 RdDesginControl";
            /// <summary>
            /// CASE_A_02 : 無 FileId
            /// </summary>
            public static string CASE_A_02_NO_FILEID = "CASE_A_02 : 無 FileId";
            /// <summary>
            /// CASE_A_03 : 不是 .dxf 檔案
            /// </summary>
            public static string CASE_A_03 = "CASE_A_03 : 不是 .dxf 檔案";
            /// <summary>
            /// CASE_A_99 : 可以開始調用 AI API
            /// </summary>
            public static string CASE_A_99 = "CASE_A_99 : 可以開始調用 AI API";


            /// <summary>
            /// CASE_B_00 : === 人工智慧研發部 ===
            /// </summary>
            public static string CASE_B_00 = "CASE_B_00 : === 人工智慧研發部 ===";
            /// <summary>
            /// CASE_B_01 : 無法訪問 AI API
            /// </summary>
            public static string CASE_B_01_NOT_ABLE_TO_VISIT_AI_API = "CASE_B_01 : 無法訪問 AI API!";
            /// <summary>
            /// CASE_B_02 : AI API 帳密錯誤
            /// </summary>
            public static string CASE_B_02_NOT_ABLE_TO_LOGIN_AI_API = "CASE_B_02 : AI API 帳密錯誤!";
            /// <summary>
            /// CASE_B_03 : AI API 無法獲得解析結果
            /// </summary>
            public static string CASE_B_03_FAILED_TO_GET_PARSE = "CASE_B_03 : AI 無法獲得解析結果!";
            /// <summary>
            /// CASE_B_04 : AI API 解析結果無法使用
            /// </summary>
            public static string CASE_B_04_PARSED_RESULT_NOT_USEFUL = "CASE_B_03 : AI 解析結果無法使用!";
            /// <summary>
            /// CASE_B_99 : 取得可用的 JSON
            /// </summary>
            public static string CASE_B_99 = "CASE_B_99 : 取得可用的 JSON";

            /// <summary>
            /// CASE_C_00 : === 途程資料管理(判斷) === 
            /// </summary>
            public static string CASE_C_00 = "CASE_C_00 : === 途程資料管理(判斷) === ";

            /// <summary>
            /// CASE_C_01 : 已有相同途程名稱 
            /// </summary>
            public static string CASE_C_01_EXISTING_ROUTINGNAME = "CASE_C_01 : 已有相同途程名稱!";

            /// <summary>
            /// CASE_C_02 : 已有相同途程(名稱不同) 
            /// </summary>
            public static string CASE_C_02_EXISTING_ROUTING_PROCESS = "CASE_C_02 : 已有相同途程(名稱不同)";
            /// <summary>
            /// CASE_C_03 : 已有相同途程(名稱不同) 
            /// </summary>
            public static string CASE_C_03_WOULD_BE_A_NEW_ROUTING = "CASE_C_03 : 確認沒有任何相同途程(不管名稱相不不同)";

            /// <summary>
            /// CASE_D_00 : === 途程資料管理(新增) === 
            /// </summary>
            public static string CASE_D_00 = "CASE_D_00 : === 途程資料管理(新增) === ";

            /// <summary>
            /// CASE_D_01 : 新增 [途程][途程製程][途程製程detial][途程品號][途程品號detail] 
            /// </summary>
            public static string CASE_D_01_TO_CREATE_ENTIRE_ROUTING = "CASE_D_01 : 新增 全套[途程] 上右左!";

            /// <summary>
            /// CASE_D_02 : 新增 [途程品號][途程品號detail]
            /// </summary>
            public static string CASE_D_02_TO_CREATE_ROUTING_ITEM_ONLY = "CASE_D_02 : 新增 半套[途程] 左! ";

            public static List<string> GetAll()
            {
                return new List<string>
                {
                    CASE_A_00,
                    CASE_A_01_RdDesignControl_NOT_FOUND,
                    CASE_A_02_NO_FILEID,
                    //CASE_A_03,
                    CASE_A_99,

                    CASE_B_00,
                    CASE_B_01_NOT_ABLE_TO_VISIT_AI_API,
                    CASE_B_02_NOT_ABLE_TO_LOGIN_AI_API,
                    CASE_B_99,

                    CASE_C_00,
                    CASE_C_01_EXISTING_ROUTINGNAME,
                    CASE_C_02_EXISTING_ROUTING_PROCESS,
                    CASE_C_03_WOULD_BE_A_NEW_ROUTING,

                    CASE_D_00,
                    CASE_D_01_TO_CREATE_ENTIRE_ROUTING,
                    CASE_D_02_TO_CREATE_ROUTING_ITEM_ONLY,

                };
            }
        }
        public class IsExistingRoutingResponse
        {
            public string RoutingName { get; set; }
            public bool IsAnySameRoutingName { get; set; }
            public MesRouting SameNameRouting { get; set; }
            public int EntireCompanyRoutingCnt { get; set; }
            public int ProcessCnt { get; set; }
            public int SameProcessCntRoutingCnt { get; set; }
            public string strGoingToAdd { get; set; }
            public bool IsAnySameProcessList { get; set; }
            public MesRouting SameProcessListDifferentRoutingName { get; set; }
            //public int RoutingId { get; set; }

        }
        public class PythonRet
        {
            public string Status { get; set; }
            public string Msg { get; set; }
            public Data Data { get; set; }
        }
        public class Data
        {
            public List<RountingProcess> RountingProcess { get; set; }
            public List<RountingItemProcess> RountingItemProcess { get; set; }
        }

        public class RountingProcess
        {
            //製程編號 製程順序 支援工程檢 工程檢頻率 是否支援包裝條碼 是否顯示在流程卡 必要過站

            // NOTE: 是否支援包裝條碼, 不在JSON中, 但是在SQL中
            public string ProcessNo { get; set; } //製程編號
            public string ProcessAlias { get; set; }//製程別名
            public int SortNumber { get; set; }//製程順序
            public string ProcessCheckStatus { get; set; } //支援工程檢
            public string ProcessCheckType { get; set; }//工程檢頻率
            public string DisplayStatus { get; set; } //是否顯示在流程卡
            public string NecessityStatus { get; set; }//必要過站
            public override int GetHashCode()
            {
                int hash = 17; // Starting value for the hash code

                // Include all properties you want to consider for the hash, except RoutingId and ProcessId
                //  hash = hash * 23 + ProcessNo.GetHashCode();
                // 不包括 ProcessNo, 因為 ProcessNo ProcessId 需要另一個開發成本
                // 最後總是還要再一一比對, 所以在這裡可以先略去
                hash = hash * 23 + SortNumber.GetHashCode();
                hash = hash * 23 + ProcessAlias.GetHashCode();
                hash = hash * 23 + DisplayStatus.GetHashCode();
                hash = hash * 23 + NecessityStatus.GetHashCode();
                hash = hash * 23 + ProcessCheckStatus.GetHashCode();
                hash = hash * 23 + (ProcessCheckType?.GetHashCode() ?? 0);
                //hash = hash * 23 + PackageFlag.GetHashCode();
                //hash = hash * 23 + Status.GetHashCode();
                //hash = hash * 23 + CreateDate.GetHashCode();
                //hash = hash * 23 + (LastModifiedDate?.GetHashCode() ?? 0);
                //hash = hash * 23 + CreateBy.GetHashCode();
                //hash = hash * 23 + (LastModifiedBy?.GetHashCode() ?? 0);

                return hash;
            }
            public List<RoutingProcessItem> RoutingProcessItem { get; set; }
        }

        public class RoutingProcessItem
        {
            public string ItemNo { get; set; }
            public string ItemName { get; set; }
            public string ItemDesc { get; set; }
            public char ChkUnique { get; set; }
        }

        public class RountingItemProcess
        {
            //"CompanyId": 2,
            //     "ProcessId": 10,
            //     "ProcessNo": "JMO-IE-0010",
            //     "ProcessName": "創成",
            //     "ProcessDesc": "基材高度0__Chu ý :Nhin Thu Nho Lam",
            //     "ProcessRemark": "",
            //     "ProcessCycleTime": "",
            //     "ProcessMoveTime": "",
            //     "ProcessSortNumber": 1,
            //     "ProcessAlias": "創成"
            public int CompanyId { get; set; }
            public int ProcessId { get; set; }
            public string ProcessNo { get; set; }
            public string ProcessName { get; set; }
            public string ProcessDesc { get; set; }
            public string ProcessRemark { get; set; }
            public string ProcessCycleTime { get; set; }//當是"", 要處理
            public string ProcessMoveTime { get; set; }
            public int ProcessSortNumber { get; set; }
            public string ProcessAlias { get; set; }
        }



        public class MesRoutingItem
        {
            //[Key]
            //[DatabaseGenerated(DatabaseGeneratedOption.Identity)]
            public int RoutingItemId { get; set; }
            public int RoutingId { get; set; }
            public int? ControlId { get; set; }
            public int MtlItemId { get; set; }
            public string Status { get; set; }
            public string RoutingItemConfirm { get; set; }
            public DateTime CreateDate { get; set; }
            public DateTime? LastModifiedDate { get; set; }
            public int CreateBy { get; set; }
            public int? LastModifiedBy { get; set; }

            // Navigation properties if needed (for foreign keys)
            // public virtual MtlItem MtlItem { get; set; }
            // public virtual RdDesignControl RdDesignControl { get; set; }
        }




        public class RdDesignControlInput
        {
            public int MtlItemId { get; set; }
            public int ControlId { get; set; }
            public int FileId { get; set; }
            //     public string FileName { get; set; }

            // ... other properties
        }


        public class Root
        {
            public string Status { get; set; }
            public PythonRet Python_Ret { get; set; }
        }

        #endregion

        #region//Get for 自動流程卡專案 ApiAiAutoRouting -- MarkChen 2023-08-22
        public List<MesRoutingItemProcess> GetMesRoutingItempProcessByRoutingItemId(int RoutingItemId)
        {
            var newList = new List<MesRoutingItemProcess>();
            var MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            using (SqlConnection conn = new SqlConnection(MainConnectionStrings))
            {
                conn.Open();

                using (SqlCommand command = new SqlCommand("SELECT * FROM [MES].[RoutingItemProcess] WHERE RoutingItemId = @RoutingItemId", conn))
                {
                    command.Parameters.AddWithValue("@RoutingItemId", RoutingItemId);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read()) // Use a while loop to read through all the rows
                        {
                            var newOne = new MesRoutingItemProcess
                            {
                                ItemProcessId = (int)reader["ItemProcessId"], //ItemProcessId 全名應該是 RoutingItemProcessId
                                RoutingItemId = (int)reader["RoutingItemId"],
                                RoutingProcessId = (int)reader["RoutingProcessId"],

                                RoutingItemProcessDesc = reader["RoutingItemProcessDesc"] == DBNull.Value ? string.Empty : reader["RoutingItemProcessDesc"].ToString(),

                                CycleTime = (int)reader["CycleTime"],
                                MoveTime = (int)reader["MoveTime"],
                                Remark = reader["Remark"] == DBNull.Value ? string.Empty : reader["Remark"].ToString(),
                                DisplayStatus = reader["DisplayStatus"] == DBNull.Value ? string.Empty : reader["DisplayStatus"].ToString(),


                                CreateDate = (DateTime)reader["CreateDate"],
                                LastModifiedDate = reader["LastModifiedDate"] as DateTime?,
                                CreateBy = (int)reader["CreateBy"],
                                LastModifiedBy = reader["LastModifiedBy"] as int?
                            };

                            newList.Add(newOne); // Add each record to the list
                        }

                        if (newList.Count == 0)
                        {
                            throw new Exception("RoutingItem not found");
                        }
                    }
                }
            }
            return newList;
        }
        public List<MesRoutingItem> GetMesRoutingItemByMtlItemId(int MtlItemId)
        {
            var newList = new List<MesRoutingItem>();
            var MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            using (SqlConnection conn = new SqlConnection(MainConnectionStrings))
            {
                conn.Open();

                using (SqlCommand command = new SqlCommand("SELECT * FROM [MES].[RoutingItem] WHERE MtlItemId = @MtlItemId", conn))
                {
                    command.Parameters.AddWithValue("@MtlItemId", MtlItemId);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read()) // Use a while loop to read through all the rows
                        {
                            var newOne = new MesRoutingItem
                            {
                                RoutingItemId = (int)reader["RoutingItemId"],
                                RoutingId = (int)reader["RoutingId"],
                                ControlId = (int)reader["ControlId"],
                                MtlItemId = (int)reader["MtlItemId"],
                                Status = reader["Status"].ToString(),
                                RoutingItemConfirm = reader["RoutingItemConfirm"].ToString(),
                                CreateDate = (DateTime)reader["CreateDate"],
                                LastModifiedDate = reader["LastModifiedDate"] as DateTime?,
                                CreateBy = (int)reader["CreateBy"],
                                LastModifiedBy = reader["LastModifiedBy"] as int?
                            };

                            newList.Add(newOne); // Add each record to the list
                        }

                        if (newList.Count == 0)
                        {
                            throw new Exception("RoutingItem not found");
                        }
                    }
                }
            }
            return newList;
        }
        public List<MesRoutingItem> GetMesRoutingItemByMtlItemIdRoutingId(int MtlItemId, int RoutingId)
        {
            var newList = new List<MesRoutingItem>();
            var MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            using (SqlConnection conn = new SqlConnection(MainConnectionStrings))
            {
                conn.Open();

                using (SqlCommand command = new SqlCommand("SELECT * FROM [MES].[RoutingItem] WHERE MtlItemId = @MtlItemId AND  RoutingId = @RoutingId", conn))
                {
                    command.Parameters.AddWithValue("@MtlItemId", MtlItemId);
                    command.Parameters.AddWithValue("@RoutingId", RoutingId);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read()) // Use a while loop to read through all the rows
                        {
                            var newOne = new MesRoutingItem
                            {
                                RoutingId = (int)reader["RoutingId"],
                                ControlId = (int)reader["ControlId"],
                                MtlItemId = (int)reader["MtlItemId"],
                                Status = reader["Status"].ToString(),
                                RoutingItemConfirm = reader["RoutingItemConfirm"].ToString(),
                                CreateDate = (DateTime)reader["CreateDate"],
                                LastModifiedDate = reader["LastModifiedDate"] as DateTime?,
                                CreateBy = (int)reader["CreateBy"],
                                LastModifiedBy = reader["LastModifiedBy"] as int?
                            };

                            newList.Add(newOne); // Add each record to the list
                        }

                        if (newList.Count == 0)
                        {
                            throw new Exception("RoutingItem not found");
                        }
                    }
                }
            }
            return newList;
        }
        public List<MesProcess> GetMesProcessByCompanyId(int CompanyId)
        {
            var newList = new List<MesProcess>();
            var MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            using (SqlConnection conn = new SqlConnection(MainConnectionStrings))
            {
                conn.Open();

                using (SqlCommand command = new SqlCommand("SELECT * FROM [MES].[Process] WHERE CompanyId = @CompanyId", conn))
                {
                    command.Parameters.AddWithValue("@CompanyId", CompanyId);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read()) // Use a while loop to read through all the rows
                        {
                            var newOne = new MesProcess
                            {
                                ProcessId = (int)reader["ProcessId"],
                                CompanyId = (int)reader["CompanyId"],
                                ProcessNo = reader["ProcessNo"].ToString(),
                                ProcessName = reader["ProcessName"].ToString(),
                                ProcessDesc = reader["ProcessDesc"].ToString(),
                                Status = reader["Status"].ToString(),
                                //RoutingConfirm = reader["RoutingConfirm"].ToString(),
                                CreateDate = (DateTime)reader["CreateDate"],
                                LastModifiedDate = reader["LastModifiedDate"] as DateTime?,
                                CreateBy = (int)reader["CreateBy"],
                                LastModifiedBy = reader["LastModifiedBy"] as int?
                            };

                            newList.Add(newOne); // Add each record to the list
                        }

                        if (newList.Count == 0)
                        {
                            throw new Exception("RoutingName not found");
                        }
                    }
                }
            }
            return newList;
        }
        public List<DboVCompanyRoutingProcessCnt> GetCompanyRoutingProcessCounts(int companyId)
        {
            List<DboVCompanyRoutingProcessCnt> result = new List<DboVCompanyRoutingProcessCnt>();

            var MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            using (SqlConnection connection = new SqlConnection(MainConnectionStrings))
            {
                connection.Open();

                string sqlQuery = @"
            SELECT CompanyId, T2.RoutingId, COUNT(*) ProcessCnt 
            FROM MES.RoutingProcess T1
            JOIN MES.Routing T2 ON T1.RoutingId = T2.RoutingId
            WHERE T2.CompanyId = @CompanyId
            GROUP BY CompanyId, T2.RoutingId";

                using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                {
                    command.Parameters.AddWithValue("@CompanyId", companyId);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(new DboVCompanyRoutingProcessCnt
                            {
                                CompanyId = (int)reader["CompanyId"],
                                RoutingId = (int)reader["RoutingId"],
                                ProcessCnt = (int)reader["ProcessCnt"]
                            });
                        }
                    }
                }
            }

            return result;
        }
        public string CombineHashCodesUsingSHA256(List<int> hashCodes)
        {
            StringBuilder combinedHashCodes = new StringBuilder();
            foreach (int hashCode in hashCodes)
            {
                combinedHashCodes.Append(hashCode.ToString());
            }

            using (SHA256 sha256Hash = SHA256.Create())
            {
                byte[] bytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(combinedHashCodes.ToString()));

                StringBuilder finalHashCode = new StringBuilder();
                for (int i = 0; i < bytes.Length; i++)
                {
                    finalHashCode.Append(bytes[i].ToString("x2"));
                }

                return finalHashCode.ToString();
            }
        }

        public List<MesRoutingProcess> GetRoutingProcessesByRoutingId(int routingId)
        {
            List<MesRoutingProcess> processes = new List<MesRoutingProcess>();

            var MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            using (SqlConnection connection = new SqlConnection(MainConnectionStrings))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("SELECT * FROM [MES].[RoutingProcess] WHERE RoutingId = @RoutingId ORDER BY SortNumber ", connection))
                {
                    command.Parameters.AddWithValue("@RoutingId", routingId);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            processes.Add(new MesRoutingProcess
                            {
                                RoutingProcessId = (int)reader["RoutingProcessId"],
                                RoutingId = (int)reader["RoutingId"],
                                ProcessId = (int)reader["ProcessId"],
                                SortNumber = (int)reader["SortNumber"],
                                ProcessAlias = reader["ProcessAlias"].ToString(),
                                DisplayStatus = reader["DisplayStatus"].ToString(),
                                NecessityStatus = reader["NecessityStatus"].ToString(),
                                ProcessCheckStatus = reader["ProcessCheckStatus"].ToString(),
                                ProcessCheckType = reader["ProcessCheckType"] as string,
                                PackageFlag = reader["PackageFlag"].ToString(),
                                Status = reader["Status"].ToString(),
                                CreateDate = (DateTime)reader["CreateDate"],
                                LastModifiedDate = reader["LastModifiedDate"] as DateTime?,
                                CreateBy = (int)reader["CreateBy"],
                                LastModifiedBy = reader["LastModifiedBy"] as int?
                            });
                        }
                    }
                }
            }

            return processes;
        }
        public MesRouting GetRoutingById(int RoutingId)
        {
            var MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            using (SqlConnection conn = new SqlConnection(MainConnectionStrings))
            {
                conn.Open();

                using (SqlCommand command = new SqlCommand("SELECT * FROM [MES].[Routing] WHERE RoutingId = @RoutingId", conn))
                {
                    command.Parameters.AddWithValue("@RoutingId", RoutingId);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return new MesRouting
                            {
                                RoutingId = (int)reader["RoutingId"],
                                CompanyId = (int)reader["CompanyId"],
                                ModeId = (int)reader["ModeId"],
                                RoutingType = reader["RoutingType"].ToString(),
                                RoutingName = reader["RoutingName"].ToString(),
                                Status = reader["Status"].ToString(),
                                RoutingConfirm = reader["RoutingConfirm"].ToString(),
                                CreateDate = (DateTime)reader["CreateDate"],
                                LastModifiedDate = reader["LastModifiedDate"] as DateTime?,
                                CreateBy = (int)reader["CreateBy"],
                                LastModifiedBy = reader["LastModifiedBy"] as int?
                            };
                        }
                        else
                        {
                            //throw new Exception("RoutingName not found");
                            return new MesRouting();
                        }
                    }
                }
            }
        }
        public IsExistingRoutingResponse IsExistingRoutingAdvAsync(Root rootObject, int int_file_id)
        {
            IsExistingRoutingResponse response = new IsExistingRoutingResponse();

            //製程編號 製程順序 製程別名 支援工程檢 工程檢頻率 是否支援包裝條碼 是否顯示在流程卡 必要過站, 記錄和內容都相同。

            ////先確認製程個數, 再以同製程個數做進一步比對
            //var lst = await GetDboVCompanyRoutingProcessCnt(2);
            ////var lst = await GetDboVCompanyRoutingProcessCnt(2);
            List<DboVCompanyRoutingProcessCnt> lst = GetCompanyRoutingProcessCounts(2);
            response.EntireCompanyRoutingCnt = lst.Count();

            //ShowTaskWithTime($"目前JMO計有途程 GetDboVCompanyRoutingProcessCnt(2)={lst.Count} 組。");
            var cntProcess = rootObject.Python_Ret.Data.RountingProcess.Count();
            response.ProcessCnt = cntProcess;
            var lst2 = lst.Where(a => a.ProcessCnt == cntProcess).ToList();
            response.SameProcessCntRoutingCnt = lst2.Count();
            //ShowTaskWithTime($"這是準備新增的", false);
            ////foreach (var x2 in rootObject.Python_Ret.Data.RountingProcess.OrderBy(a => a.SortNumber).ToList())
            ////{
            ////    ShowTaskWithTime($"SortName= {x2.SortNumber} HashCode={x2.GetHashCode()}", false);
            ////}

            List<int> goingToAdd = rootObject.Python_Ret.Data.RountingProcess
                        .OrderBy(a => a.SortNumber)
                        .Select(a => a.GetHashCode())
                        .ToList();
            string strGoingToAdd = CombineHashCodesUsingSHA256(goingToAdd);
            response.strGoingToAdd = strGoingToAdd;


            //ShowTaskWithTime($"其 SHA256 為 {strGoingToAdd}", false);

            //ShowTaskWithTime($"剛解析的途程其 RoutingProcess 計有 {cntProcess} 個。", false);

            //ShowTaskWithTime($"其中JMO途程 有RoutingProcess {cntProcess} 個, 計有 {lst.Count} 組。", false);
            //ShowTaskWithTime($"=== 使用密碼學原理做有效率的比對 ===", false);
            //ShowTaskWithTime($"將 RoutingId 列出其所有 製程。", false);
            var rec = 0;
            foreach (var x in lst.OrderByDescending(a => a.RoutingId))
            {
                rec++;
                //    //ShowTaskWithTime($"第 {rec} 組, 列出其所有 製程。", false);
                //    var lstProcess = await aiContext.MesRoutingProcesss.Where(a => a.RoutingId == x.RoutingId).OrderBy(a => a.SortNumber).ToListAsync();
                var lstProcess = GetRoutingProcessesByRoutingId(x.RoutingId);

                List<int> known = lstProcess.Select(a => a.GetHashCode()).ToList();
                string strKnow = CombineHashCodesUsingSHA256(known);
                //    //ShowTaskWithTime($"第 {rec} 組, 其 SHA256 為 {strKnow}", false);
                if (strGoingToAdd == strKnow)
                {
                    //        //ShowTaskWithTime($"    **** 找到符合的 RoutingId={x.RoutingId}。再進一步核對第三層!!!", false);
                    //        ShowTaskWithTime($"    **** 找到符合的 RoutingId={x.RoutingId}。視同途程相同!!!", false);

                    MesRouting known2 = GetRoutingById(x.RoutingId);
                    response.IsAnySameProcessList = true;
                    response.SameProcessListDifferentRoutingName = known2;
                    return response;
                }
                //        //ShowTaskWithTime($"    RoutingName is {known2.RoutingName}", false);

                //        //ShowTaskWithTime($"\n\n  開發測試不同情境, 可以 刻意動一下[途程製程]的順序, 或任意製程別名   \n\n", false);
                //        //ShowTaskWithTime($"\n\n  開發測試不同情境, 可以 刪掉途程 -del {x.RoutingId}  \n\n", false);
                //        //ShowTaskWithTime($"\n\n  要注意  ALTER TABLE [MES].[RoutingItemProcess]  WITH CHECK ADD  CONSTRAINT [FK2_RoutingItemProcess] FOREIGN KEY([RoutingProcessId])\r\nREFERENCES [MES].[RoutingProcess] ([RoutingProcessId]) \n\n", false);
                //        ShowTaskWithTime($"=== 使用密碼學原理做有效率的比對 === 發現第一個 相同的 途程", false);
                //        return (true, x.RoutingId);
                //    }


            }
            //ShowTaskWithTime($"=== 使用密碼學原理做有效率的比對 === 沒有發現任何相同的 途程", false);

            ////FOR DEBUG ONLY, 這樣子, 就不會新增測試中的途程, 還要再刪除才能再測試
            ////先以 SHA256 比到第二層
            //return (false, 0);
            return response;
        }
        public MesRouting GetRoutingByName(string routingName)
        {
            var MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            using (SqlConnection conn = new SqlConnection(MainConnectionStrings))
            {
                conn.Open();

                using (SqlCommand command = new SqlCommand("SELECT * FROM [MES].[Routing] WHERE RoutingName = @RoutingName", conn))
                {
                    command.Parameters.AddWithValue("@RoutingName", routingName);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return new MesRouting
                            {
                                RoutingId = (int)reader["RoutingId"],
                                CompanyId = (int)reader["CompanyId"],
                                ModeId = (int)reader["ModeId"],
                                RoutingType = reader["RoutingType"].ToString(),
                                RoutingName = reader["RoutingName"].ToString(),
                                Status = reader["Status"].ToString(),
                                RoutingConfirm = reader["RoutingConfirm"].ToString(),
                                CreateDate = (DateTime)reader["CreateDate"],
                                LastModifiedDate = reader["LastModifiedDate"] as DateTime?,
                                CreateBy = (int)reader["CreateBy"],
                                LastModifiedBy = reader["LastModifiedBy"] as int?
                            };
                        }
                        else
                        {
                            //throw new Exception("RoutingName not found");
                            return new MesRouting();
                        }
                    }
                }
            }
        }
        public RdDesignControl GetRdDesignControl(int MtlItemId, int ControlId)
        {
            var MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            using (SqlConnection conn = new SqlConnection(MainConnectionStrings))
            {
                conn.Open();

                //string query = @"
                //    SELECT T2.MtlItemId, T1.*
                //    FROM [PDM].[RdDesignControl] AS T1
                //    JOIN [PDM].[RdDesign] AS T2 ON T1.DesignId = T2.DesignId
                //    WHERE T2.MtlItemId = @MtlItemId 
                //    AND T1.ControlId = @ControlId;
                //    ";
                string query = @"
                SELECT T3.MtlItemNo, T3.MtlItemName, T2.MtlItemId, T1.*
                FROM [PDM].[RdDesignControl] AS T1
                JOIN [PDM].[RdDesign] AS T2 ON T1.DesignId = T2.DesignId
                JOIN [PDM].[MtlItem] AS T3 ON T2.MtlItemId = T3.MtlItemId
                WHERE T2.MtlItemId = @MtlItemId 
                AND T1.ControlId = @ControlId;
                ";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@MtlItemId", MtlItemId);
                    cmd.Parameters.AddWithValue("@ControlId", ControlId);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return new RdDesignControl
                            {
                                MtlItemNo = reader["MtlItemNo"].ToString(),
                                MtlItemName = reader["MtlItemName"].ToString(),

                                MtlItemId = (int)reader["MtlItemId"],
                                ControlId = (int)reader["ControlId"],
                                DesignId = (int)reader["DesignId"],
                                //Cad2DFile = (int)reader["Cad2DFile"],
                                Cad2DFile = reader["Cad2DFile"] != DBNull.Value ? (int)reader["Cad2DFile"] : (int?)null,
                                // ... other properties
                            };
                        }
                    }
                }
            }

            return null; // Return null if no matching record is found
        }

        #endregion

        #endregion

        #region //Add

        //using System.Data.SqlClient;
        #region 新增整套 途程
        public RdDesignControlResult ExecuteTransaction_CASE_D_01(RdDesignControlResult caseObj,
            int MtlItemId, int ControlId)
        {
            MesRouting level1 = null;
            var dtNow = DateTime.Now;
            var strNow = dtNow.ToString("yyyy-MM-dd HH:mm:ss");
            var intUser = 0;

            caseObj.Step4.CaseType = " 新增 全套途程, 上方[途程], 右邊[途程製程][途程製程屬性] 左邊[途程品號][加工細節] ";

            caseObj.Step4.DebugMsg = " 不能 by ref, 必需 by Val " + caseObj.Step4.DebugMsg;
            string connectionString = ConfigurationManager.AppSettings["MainDb"];
            List<MesProcess> processList = GetMesProcessByCompanyId(2);

            // FOR DEBUG TO VERIFY THE LIST
            //caseObj.CASE_D_Data.ProcessList = processList; 

            // FOR DEBUG TO VERIFY THE connectionString
            //caseObj.CASE_D_Data.ConnStr = connectionString;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Start a local transaction.
                SqlTransaction sqlTransaction = connection.BeginTransaction();
                caseObj.Step4.TransactionStart = DateTime.Now;

                using (SqlCommand command = connection.CreateCommand())
                {
                    command.Transaction = sqlTransaction;
                    command.CommandTimeout = 90; // Timeout in seconds 之前 30就跳掉了
                    try
                    {
                        var sql1 = $@"
                    INSERT INTO [MES].[Routing]
                           ([CompanyId]
                           ,[ModeId]
                           ,[RoutingType]
                           ,[RoutingName]
                           ,[Status]
                           ,[RoutingConfirm]
                           ,[CreateDate]
                           ,[LastModifiedDate]
                           ,[CreateBy]
                           ,[LastModifiedBy])
                     VALUES
                           (2
                           ,1
                           ,'1'
                           ,'{caseObj.Step3.RoutingName}'
                           ,'A'
                           ,'Y'
                           ,'{strNow}'
                           ,'{strNow}'
                           ,{intUser}
                           ,{intUser}
                            )
                            ; SELECT SCOPE_IDENTITY() AS NewID
                     ";
                        caseObj.Step4.DebugSQL1 = sql1.Replace("\r\n", " "); ;
                        // Execute your commands here.
                        command.CommandText = sql1;
                        //command.ExecuteNonQuery();
                        // Execute the INSERT command and retrieve the new ID.
                        object newID = command.ExecuteScalar();
                        int level1RoutingId = Convert.ToInt32(newID);

                        //MesRouting level1 = GetRoutingByName(caseObj.CASE_C_Data.RoutingName);
                        caseObj.Step4.Level1RoutingId = level1RoutingId;


                        //

                        var sql4 = $@"
                        INSERT INTO [MES].[RoutingItem]
                                   ([RoutingId]
                                   ,[ControlId]
                                   ,[MtlItemId]
                                   ,[Status]
                                   ,[RoutingItemConfirm]
                                   ,[CreateDate]
                                   ,[LastModifiedDate]
                                   ,[CreateBy]
                                   ,[LastModifiedBy])
                             VALUES
                                   ({ caseObj.Step4.Level1RoutingId}
                                   ,{ControlId}
                                   ,{MtlItemId}
                                   ,'A'
                                   ,'Y'
                                  ,GETDATE()
                                  ,GETDATE()
                                   ,0
                                   ,0)   ; SELECT SCOPE_IDENTITY() AS NewID";
                        command.CommandText = sql4;
                        caseObj.Step4.DebugSQL4 = sql4.Replace("\r\n", " ");

                        object newLevel4ID = command.ExecuteScalar();
                        int level4RoutingItemId = Convert.ToInt32(newLevel4ID);

                        //MesRouting level1 = GetRoutingByName(caseObj.CASE_C_Data.RoutingName);
                        caseObj.Step4.Level4RoutingItemId = level4RoutingItemId;



                        foreach (var x in caseObj.Step2.Python_Ret.Data.RountingProcess)
                        {

                            var objProcess = processList.Where(a => a.ProcessNo == x.ProcessNo).FirstOrDefault();
                            if (objProcess == null)
                            {
                                throw new Exception("ProcessNo is not on the current system!");
                            }
                            //var processId = objProcess.ProcessId;
                            caseObj.Step4.ProcessId = objProcess.ProcessId;
                            caseObj.Step4.ProcessNo = objProcess.ProcessNo;

                            var sql2 = $@"
                                INSERT INTO [MES].[RoutingProcess]
                                           ([RoutingId]
                                           ,[ProcessId]
                                           ,[SortNumber]
                                           ,[ProcessAlias]
                                           ,[DisplayStatus]
                                           ,[NecessityStatus]
                                           ,[ProcessCheckStatus]
                                           ,[ProcessCheckType]
                                           ,[PackageFlag]
                                           ,[Status]
                                           ,[CreateDate]
                                           ,[LastModifiedDate]
                                           ,[CreateBy]
                                           ,[LastModifiedBy])
                                VALUES(
                                    {caseObj.Step4.Level1RoutingId}
                                    ,{caseObj.Step4.ProcessId}
                                    ,{x.SortNumber}
                                    ,'{x.ProcessAlias}'
                                     ,'{x.DisplayStatus}'
                                    ,'{x.NecessityStatus}'
                                    ,'{x.ProcessCheckStatus}'
                                    ,'{x.ProcessCheckType}'
                        
                                    ,'N' --<PackageFlag, nvarchar(2),>
                                    ,'A' -- <Status, nvarchar(2),>
                                    ,'{strNow}'
                                    ,'{strNow}'
                                    ,{intUser}
                                    ,{intUser}
                                )
                                ; SELECT SCOPE_IDENTITY() AS newLevel2ID
                                ";
                            //WHY sql2 is a problem, to be NULL???
                            caseObj.Step4.DebugSQL2 = sql2.Replace("\r\n", " ");

                            command.CommandText = sql2;
                            // Execute the INSERT command and retrieve the new ID for the second level.
                            object newLevel2ID = command.ExecuteScalar();
                            if (newLevel2ID == null)
                            {
                                throw new Exception("第二層沒有 能夠寫入");
                            }
                            int level2RoutingProcessId = Convert.ToInt32(newLevel2ID);
                            caseObj.Step4.Level2RoutingProcessId = level2RoutingProcessId;

                            //目前第三層只有刻字
                            var objlevel3 = caseObj.Step2.Python_Ret.Data.RountingProcess.Where(a => a.ProcessNo == caseObj.Step4.ProcessNo).FirstOrDefault();
                            if (objlevel3.RoutingProcessItem != null)
                            {
                                foreach (var x3 in objlevel3.RoutingProcessItem)
                                {
                                    var sql3 = $@"
                                        INSERT INTO [MES].[RoutingProcessItem]
                                                   ([RoutingProcessId]
                                                   ,[ItemNo]
                                                   ,[ChkUnique]
                                                   ,[CreateDate]
                                                   ,[LastModifiedDate]
                                                   ,[CreateBy]
                                                   ,[LastModifiedBy])
                                             VALUES
                                                   ({ caseObj.Step4.Level2RoutingProcessId}
                                                    ,'{x3.ItemNo}'
                                                   ,'{x3.ChkUnique}'
                                                   ,'{strNow}'
                                                   ,'{strNow}'
                                                   ,{intUser}
                                                   ,{intUser})
                                            ; SELECT SCOPE_IDENTITY() AS newLevel3ID
                                        ";

                                    caseObj.Step4.DebugSQL3 = sql3.Replace("\r\n", " ");

                                    command.CommandText = sql3;
                                    object newLevel3ID = command.ExecuteScalar();
                                    caseObj.Step4.Level3RoutingProcessItemId = Convert.ToInt32(newLevel3ID);


                                }
                            }


                            //LEVEL 5, 就是左邊的第二層
                            var objLevel5 = caseObj.Step2.Python_Ret.Data.RountingItemProcess
                                .Where(a => a.ProcessNo == caseObj.Step4.ProcessNo).FirstOrDefault();
                            if (objLevel5 != null)
                            {
                                decimal processCycleTime = 0; // Default value
                                if (!string.IsNullOrEmpty(objLevel5.ProcessCycleTime))
                                {
                                    if (!decimal.TryParse(objLevel5.ProcessCycleTime, out processCycleTime))
                                    {
                                        // Handle parsing error if needed
                                    }
                                }

                                decimal processMoveTime = 0; // Default value
                                if (!string.IsNullOrEmpty(objLevel5.ProcessMoveTime))
                                {
                                    if (!decimal.TryParse(objLevel5.ProcessMoveTime, out processMoveTime))
                                    {
                                        // Handle parsing error if needed
                                    }
                                }


                                var sql5 = $@"

                                INSERT INTO [MES].[RoutingItemProcess]
                                           ([RoutingItemId]
                                           ,[RoutingProcessId]
                                           ,[RoutingItemProcessDesc]
                                           ,[CycleTime]
                                           ,[MoveTime]
                                           ,[Remark]
                                           ,[DisplayStatus]
                                         --  ,[AttrSetting]
                                           ,[CreateDate]
                                           ,[LastModifiedDate]
                                           ,[CreateBy]
                                           ,[LastModifiedBy])
                                     VALUES
                                           ({ caseObj.Step4.Level4RoutingItemId}
                                           ,{ caseObj.Step4.Level2RoutingProcessId}
                                           ,'{objLevel5.ProcessDesc}'
                                           ,{processCycleTime}
                                           ,{processMoveTime}
                                           ,'{objLevel5.ProcessRemark}'
                                           ,'Y' --<DisplayStatus, nvarchar(2),>
                                           --,<AttrSetting, nvarchar(200),>
                                           ,'{strNow}'
                                           ,'{strNow}'
                                           ,{intUser}
                                           ,{intUser}        
                                            )
                                            ; SELECT SCOPE_IDENTITY() AS newLevel5ID
                                        ";
                                caseObj.Step4.DebugSQL5 = sql5.Replace("\r\n", " ");

                                command.CommandText = sql5;
                                object newLevel5ID = command.ExecuteScalar();
                                if (newLevel5ID == null)
                                {
                                    throw new Exception("第5層沒有 能夠寫入");
                                }
                                int level5RoutingItemProcessId = Convert.ToInt32(newLevel5ID);
                                caseObj.Step4.Level5RoutingItemProcessId = level5RoutingItemProcessId;
                            }
                        }

                        //command.CommandText = "INSERT INTO Table2 (Column2) VALUES ('Value2')";
                        //command.ExecuteNonQuery();

                        // Commit the transaction.
                        sqlTransaction.Commit();

                        caseObj.Step4.TransactionEnd = DateTime.Now;
                        caseObj.Step4.TransactionResult = "Transaction committed.";

                        Console.WriteLine("Transaction committed.");
                    }

                    catch (Exception ex)
                    {
                        // Attempt to roll back the transaction.
                        try
                        {
                            sqlTransaction.Rollback();
                            caseObj.Step4.TransactionEnd = DateTime.Now;
                            caseObj.Step4.TransactionResult = $"{ex.Message}  Transaction Rollback.";
                        }
                        catch (Exception exRollback)
                        {
                            // Handle any errors that may have occurred on the server when rolling back the transaction.
                            Console.WriteLine($"Rollback Exception Type: {exRollback.GetType()}");
                            Console.WriteLine($"Rollback Exception: {exRollback.Message}");
                            caseObj.Step4.TransactionResult = "Transaction Rollback. exception  {exRollback.Message}";
                        }

                        Console.WriteLine($"Exception Type: {ex.GetType()}");
                        Console.WriteLine($"Exception: {ex.Message}");
                    }
                }
            }

            // SOP to return 
            // NOTE: not able to use by ref 
            // 無法呼叫控制器 'Business_Manager.Controllers.RoutingController' 上的動作方法 'Void
            // ExecuteTransaction_CASE_D_01(RdDesignControlResult ByRef)'，因為參數 'RdDesignControlResult & amp; caseObj'
            // 是以傳址方式傳遞。
            return caseObj;
        }
        #endregion

        #region 新增整套 途程, as per meeting 09/01 mgr's advice, reviewed by Mark, 203-09-06
        public RdDesignControlResult ExecuteTransaction_Meeting0901(RdDesignControlResult caseObj,
            int MtlItemId, int ControlId)
        {
            //   MesRouting level1 = null;
            var dtNow = DateTime.Now;
            var strNow = dtNow.ToString("yyyy-MM-dd HH:mm:ss");
            var intUser = 0; //以0做為後續要顯示用戶時為AI

            //caseObj.Step4.CaseType = " as per meeting 09/01 mgr's advice 新增 全套途程, 上方[途程], 右邊[途程製程][途程製程屬性] 左邊[途程品號][加工細節] ";
            //原本想 by ref, 在此有限制, 沒關係, 就以 obj 為參數並回傳, 過程蒐集到不同階段的資訊
            //caseObj.Step4.DebugMsg = " 不能 by ref, 必需 by Val " + caseObj.Step4.DebugMsg;

            string connectionString = ConfigurationManager.AppSettings["MainDb"];
            List<MesProcess> processList = GetMesProcessByCompanyId(2); //目前只限JMO, 後續如要推廣到其它公司, 要帶入公司別ID

            //這是以 人工智慧研發部 解析後的JSON擬的途程名稱
            caseObj.Step3.RoutingName = GetRoutingNameV4(caseObj.listAiProcess, caseObj.Z_FilelId);

            // FOR DEBUG TO VERIFY THE LIST
            //caseObj.CASE_D_Data.ProcessList = processList; 

            // FOR DEBUG TO VERIFY THE connectionString
            //caseObj.CASE_D_Data.ConnStr = connectionString;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Start a local transaction.
                SqlTransaction sqlTransaction = connection.BeginTransaction();
                caseObj.Step4.TransactionStart = DateTime.Now;

                using (SqlCommand command = connection.CreateCommand())
                {
                    command.Transaction = sqlTransaction;
                    command.CommandTimeout = 90; // Timeout in seconds 之前 30就跳掉了? HOW To recreate the situation that triggers the error 
                    try
                    {
                        #region // 上
                        var sql1 = $@"
                        INSERT INTO [MES].[Routing]
                           ([CompanyId]
                           ,[ModeId]
                           ,[RoutingType]
                           ,[RoutingName]
                           ,[Status]
                           ,[RoutingConfirm]
                           ,[CreateDate]
                           ,[LastModifiedDate]
                           ,[CreateBy]
                           ,[LastModifiedBy])
                         VALUES
                           (2
                           ,1
                           ,'1'
                           ,'{caseObj.Step3.RoutingName}'
                           ,'A'
                           ,'Y'
                           ,'{strNow}'
                           ,'{strNow}'
                           ,{intUser}
                           ,{intUser}
                            )
                            ; SELECT SCOPE_IDENTITY() AS NewID
                         ";
                        caseObj.Step4.DebugSQL1 = sql1.Replace("\r\n", " "); ;
                        // Execute your commands here.
                        command.CommandText = sql1;
                        //command.ExecuteNonQuery();
                        // Execute the INSERT command and retrieve the new ID.
                        object newID = command.ExecuteScalar();
                        int level1RoutingId = Convert.ToInt32(newID);

                        //MesRouting level1 = GetRoutingByName(caseObj.CASE_C_Data.RoutingName);
                        caseObj.Step4.Level1RoutingId = level1RoutingId;
                        #endregion

                        #region // 左1

                        var sql4 = $@"
                        INSERT INTO [MES].[RoutingItem]
                                   ([RoutingId]
                                   ,[ControlId]
                                   ,[MtlItemId]
                                   ,[Status]
                                   ,[RoutingItemConfirm]
                                   ,[CreateDate]
                                   ,[LastModifiedDate]
                                   ,[CreateBy]
                                   ,[LastModifiedBy])
                             VALUES
                                   ({ caseObj.Step4.Level1RoutingId}
                                   ,{ControlId}
                                   ,{MtlItemId}
                                   ,'A'
                                   ,'Y'
                                  ,GETDATE()
                                  ,GETDATE()
                                   ,0
                                   ,0)   ; SELECT SCOPE_IDENTITY() AS NewID";
                        command.CommandText = sql4;
                        caseObj.Step4.DebugSQL4 = sql4.Replace("\r\n", " ");

                        object newLevel4ID = command.ExecuteScalar();
                        int level4RoutingItemId = Convert.ToInt32(newLevel4ID);

                        //MesRouting level1 = GetRoutingByName(caseObj.CASE_C_Data.RoutingName);
                        caseObj.Step4.Level4RoutingItemId = level4RoutingItemId;
                        #endregion

                        foreach (var x in caseObj.listAiProcess)
                        {
                            #region // 右1
                            var objProcess = processList.Where(a => a.ProcessNo == x.ProcessNo).FirstOrDefault();
                            if (objProcess == null)
                            {
                                throw new Exception("ProcessNo is not on the current system!");
                            }
                            caseObj.Step4.ProcessId = objProcess.ProcessId;
                            caseObj.Step4.ProcessNo = objProcess.ProcessNo;

                            var displayStatusStr = x.DisplayStatus ? "Y" : "N";
                            var necessityStatusStr = x.NecessityStatus ? "Y" : "N";

                            var sql2 = $@"
                                INSERT INTO [MES].[RoutingProcess]
                                    ([RoutingId]
                                    ,[ProcessId]
                                    ,[SortNumber]
                                    ,[ProcessAlias]
                                    ,[DisplayStatus]
                                    ,[NecessityStatus]
                                    ,[ProcessCheckStatus]
                                    ,[ProcessCheckType]
                                    ,[PackageFlag]
                                    ,[Status]
                                    ,[CreateDate]
                                    ,[LastModifiedDate]
                                    ,[CreateBy]
                                    ,[LastModifiedBy])
                                VALUES(
                                    {caseObj.Step4.Level1RoutingId}
                                    ,{caseObj.Step4.ProcessId}
                                    ,{x.SortNumber}
                                    ,'{x.ProcessAlias}'
                                    ,'{displayStatusStr}'  -- Converted to 'Y' or 'N'
                                    ,'{necessityStatusStr}' -- Converted to 'Y' or 'N'
                                    ,'{x.ProcessCheckStatus}'
                                    ,'{x.ProcessCheckType}'
                                    ,'N' --<PackageFlag, nvarchar(2),>
                                    ,'A' -- <Status, nvarchar(2),>
                                    ,'{strNow}'
                                    ,'{strNow}'
                                    ,{intUser}
                                    ,{intUser}
                                )
                                ; SELECT SCOPE_IDENTITY() AS newLevel2ID
                                ";

                            // for Debug purpose
                            caseObj.Step4.DebugSQL2 = sql2.Replace("\r\n", " ");
                            command.CommandText = sql2;

                            // Execute the INSERT command and retrieve the new ID for the second level.
                            object newLevel2ID = command.ExecuteScalar();
                            if (newLevel2ID == null)
                            {
                                throw new Exception("第二層沒有 能夠寫入");
                            }
                            int level2RoutingProcessId = Convert.ToInt32(newLevel2ID);
                            caseObj.Step4.Level2RoutingProcessId = level2RoutingProcessId;
                            #endregion

                            #region // 右2,先略過。因為目前沒有UI給用戶微調 AI JSON 的製程屬性。目前已知只有刻字有記錄。
                            /*
                            var objlevel3 = caseObj.Step2.Python_Ret.Data.RountingProcess.Where(a => a.ProcessNo == caseObj.Step4.ProcessNo).FirstOrDefault();
                            if (objlevel3.RoutingProcessItem != null)
                            {
                                foreach (var x3 in objlevel3.RoutingProcessItem)
                                {
                                    var sql3 = $@"
                                        INSERT INTO [MES].[RoutingProcessItem]
                                                   ([RoutingProcessId]
                                                   ,[ItemNo]
                                                   ,[ChkUnique]
                                                   ,[CreateDate]
                                                   ,[LastModifiedDate]
                                                   ,[CreateBy]
                                                   ,[LastModifiedBy])
                                             VALUES
                                                   ({ caseObj.Step4.Level2RoutingProcessId}
                                                    ,'{x3.ItemNo}'
                                                   ,'{x3.ChkUnique}'
                                                   ,'{strNow}'
                                                   ,'{strNow}'
                                                   ,{intUser}
                                                   ,{intUser})
                                            ; SELECT SCOPE_IDENTITY() AS newLevel3ID
                                        ";

                                    caseObj.Step4.DebugSQL3 = sql3.Replace("\r\n", " ");

                                    command.CommandText = sql3;
                                    object newLevel3ID = command.ExecuteScalar();
                                    caseObj.Step4.Level3RoutingProcessItemId = Convert.ToInt32(newLevel3ID);
                                    

                                }
                            }
                            */
                            #endregion

                            #region // 左2,先略過。因為目前沒有UI給用戶微調 AI JSON 途程品號 的 加工細節。


                            //LEVEL 5, 就是左邊的第二層
                            /*
                            var objLevel5 = caseObj.Step2.Python_Ret.Data.RountingItemProcess
                                .Where(a => a.ProcessNo == caseObj.Step4.ProcessNo).FirstOrDefault();
                            if (objLevel5 != null)
                            {
                                decimal processCycleTime = 0; // Default value
                                if (!string.IsNullOrEmpty(objLevel5.ProcessCycleTime))
                                {
                                    if (!decimal.TryParse(objLevel5.ProcessCycleTime, out processCycleTime))
                                    {
                                        // Handle parsing error if needed
                                    }
                                }

                                decimal processMoveTime = 0; // Default value
                                if (!string.IsNullOrEmpty(objLevel5.ProcessMoveTime))
                                {
                                    if (!decimal.TryParse(objLevel5.ProcessMoveTime, out processMoveTime))
                                    {
                                        // Handle parsing error if needed
                                    }
                                }
                                
                                var sql5 = $@"
                                INSERT INTO [MES].[RoutingItemProcess]
                                           ([RoutingItemId]
                                           ,[RoutingProcessId]
                                           ,[RoutingItemProcessDesc]
                                           ,[CycleTime]
                                           ,[MoveTime]
                                           ,[Remark]
                                           ,[DisplayStatus]
                                         --  ,[AttrSetting]
                                           ,[CreateDate]
                                           ,[LastModifiedDate]
                                           ,[CreateBy]
                                           ,[LastModifiedBy])
                                     VALUES
                                           ({ caseObj.Step4.Level4RoutingItemId}
                                           ,{ caseObj.Step4.Level2RoutingProcessId}
                                           ,'{objLevel5.ProcessDesc}'
                                           ,{processCycleTime}
                                           ,{processMoveTime}
                                           ,'{objLevel5.ProcessRemark}'
                                           ,'Y' --<DisplayStatus, nvarchar(2),>
                                           --,<AttrSetting, nvarchar(200),>
                                           ,'{strNow}'
                                           ,'{strNow}'
                                           ,{intUser}
                                           ,{intUser}        
                                            )
                                            ; SELECT SCOPE_IDENTITY() AS newLevel5ID
                                        ";
                                caseObj.Step4.DebugSQL5 = sql5.Replace("\r\n", " ");

                                command.CommandText = sql5;
                                object newLevel5ID = command.ExecuteScalar();
                                if (newLevel5ID == null)
                                {
                                    throw new Exception("第5層沒有 能夠寫入");
                                }
                                int level5RoutingItemProcessId = Convert.ToInt32(newLevel5ID);
                                caseObj.Step4.Level5RoutingItemProcessId = level5RoutingItemProcessId;
                           
                            }
                             */
                            #endregion
                        }

                        //command.CommandText = "INSERT INTO Table2 (Column2) VALUES ('Value2')";
                        //command.ExecuteNonQuery();

                        // Commit the transaction.
                        sqlTransaction.Commit();
                        caseObj.Status = "成功!";
                        caseObj.Debug = "已生成 途程 " + caseObj.Step3.RoutingName + " 。按目前要求, 不含原本的 右2 及 左2, 因為沒有 UI 供用戶微調。";
                        caseObj.Step4.TransactionEnd = DateTime.Now;
                        caseObj.Step4.TransactionResult = "Transaction committed.";

                        Console.WriteLine("Transaction committed.");
                    }

                    catch (Exception ex)
                    {
                        // Attempt to roll back the transaction.
                        try
                        {
                            sqlTransaction.Rollback();
                            caseObj.Step4.TransactionEnd = DateTime.Now;
                            caseObj.Step4.TransactionResult = $"{ex.Message}  Transaction Rollback.";
                            caseObj.Status = "異常!";
                            caseObj.Debug = $"{ex.Message}  Transaction Rollback.";
                        }
                        catch (Exception exRollback)
                        {
                            // Handle any errors that may have occurred on the server when rolling back the transaction.
                            Console.WriteLine($"Rollback Exception Type: {exRollback.GetType()}");
                            Console.WriteLine($"Rollback Exception: {exRollback.Message}");
                            caseObj.Step4.TransactionResult = "Transaction Rollback. exception  {exRollback.Message}";
                            caseObj.Status = "異常!";
                            caseObj.Debug = $"{exRollback.Message}  ";
                        }
                        Console.WriteLine($"Exception Type: {ex.GetType()}");
                        Console.WriteLine($"Exception: {ex.Message}");
                    }
                }
            }
            // 將原參數,蒐集到過程資訊後回傳
            return caseObj;
        }
        #endregion


        #region 新增 半套 途程, 即左邊的[途程品號]及[加工細節]
        public RdDesignControlResult ExecuteTransaction_CASE_D_02(RdDesignControlResult caseObj,
            int MtlItemId, int ControlId)
        {
            MesRouting level1 = null;
            var dtNow = DateTime.Now;
            var strNow = dtNow.ToString("yyyy-MM-dd HH:mm:ss");
            var intUser = 0;
            caseObj.Step4.CaseType = " 新增 半套途程, 即左邊的[途程品號]及[加工細節], 先找到之前己有的當成模板... ";

            //return caseObj;

            caseObj.Step4.DebugMsg = " 新增 半套 途程, 即左邊的[途程品號]及[加工細節] , 使用上一動已知的 MesRouting";

            string connectionString = ConfigurationManager.AppSettings["MainDb"];
            List<MesProcess> processList = GetMesProcessByCompanyId(2);



            caseObj.Step4.Level1RoutingId = caseObj.Step3.SameProcessListDifferentRoutingName.RoutingId;

            //List<MesRoutingProcess> routingProcessList = GetRoutingProcessesByRoutingId(caseObj.CASE_D_Data.Level1RoutingId);
            caseObj.Step4.RoutingProcessList = GetRoutingProcessesByRoutingId(caseObj.Step4.Level1RoutingId);
            List<MesRoutingItem> routingItemList = GetMesRoutingItemByMtlItemId(MtlItemId);
            // need to limit to same RoutingId as well
            var objTemplate = routingItemList.Where(a => a.RoutingId == caseObj.Step4.Level1RoutingId).LastOrDefault();
            if (objTemplate == null)
            {
                throw new Exception("ERR objTemplate=routingItemList.Where(a => a.RoutingId == caseObj.CASE_D_Data.Level1RoutingId).LastOrDefault();");
            }
            caseObj.Step4.RoutingId = objTemplate.RoutingId;

            caseObj.Step4.TemplateRoutingItemId = objTemplate.RoutingItemId;// WHY 0???

            //List<MesItem
            List<MesRoutingItemProcess> listX = GetMesRoutingItempProcessByRoutingItemId(caseObj.Step4.TemplateRoutingItemId);
            caseObj.Step4.TemplateRoutingItemProcess = listX;
            //return caseObj;

            if (caseObj.Step4.RoutingProcessList == null)
            {
                throw new Exception("caseObj.CASE_D_Data.RoutingProcessList is not supposed to be null");
            }

            // FOR DEBUG TO VERIFY THE LIST
            //caseObj.CASE_D_Data.ProcessList = processList; 

            // FOR DEBUG TO VERIFY THE connectionString
            //caseObj.CASE_D_Data.ConnStr = connectionString;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Start a local transaction.
                SqlTransaction sqlTransaction = connection.BeginTransaction();
                caseObj.Step4.TransactionStart = DateTime.Now;

                using (SqlCommand command = connection.CreateCommand())
                {
                    command.Transaction = sqlTransaction;
                    command.CommandTimeout = 90; // Timeout in seconds 之前 30就跳掉了
                    try
                    {
                        //上方[途程]已有

                        //    var sql1 = $@"
                        //INSERT INTO [MES].[Routing]
                        //       ([CompanyId]
                        //       ,[ModeId]
                        //       ,[RoutingType]
                        //       ,[RoutingName]
                        //       ,[Status]
                        //       ,[RoutingConfirm]
                        //       ,[CreateDate]
                        //       ,[LastModifiedDate]
                        //       ,[CreateBy]
                        //       ,[LastModifiedBy])
                        // VALUES
                        //       (2
                        //       ,1
                        //       ,'1'
                        //       ,'{caseObj.CASE_C_Data.RoutingName}'
                        //       ,'A'
                        //       ,'Y'
                        //       ,'{strNow}'
                        //       ,'{strNow}'
                        //       ,{intUser}
                        //       ,{intUser}
                        //        )
                        //        ; SELECT SCOPE_IDENTITY() AS NewID
                        // ";
                        //caseObj.CASE_D_Data.DebugSQL1 = sql1.Replace("\r\n", " "); ;
                        //command.CommandText = sql1;
                        //object newID = command.ExecuteScalar();
                        //int level1RoutingId = Convert.ToInt32(newID);

                        // RoutingItemId    RoutingId   ControlId   MtlItemId
                        // 6506             4375        3987        264065
                        // 6514             4375        3985        264065

                        //左邊[途程品號]

                        var sql4 = $@"
                        INSERT INTO [MES].[RoutingItem]
                                   ([RoutingId]
                                   ,[ControlId]
                                   ,[MtlItemId]
                                   ,[Status]
                                   ,[RoutingItemConfirm]
                                   ,[CreateDate]
                                   ,[LastModifiedDate]
                                   ,[CreateBy]
                                   ,[LastModifiedBy])
                             VALUES
                                   ({ caseObj.Step4.Level1RoutingId}
                                   ,{ControlId}
                                   ,{MtlItemId}
                                   ,'A'
                                   ,'Y'
                                  ,GETDATE()
                                  ,GETDATE()
                                   ,0
                                   ,0)   ; SELECT SCOPE_IDENTITY() AS NewID";
                        command.CommandText = sql4;
                        caseObj.Step4.DebugSQL4 = sql4.Replace("\r\n", " ");

                        object newLevel4ID = command.ExecuteScalar();
                        int level4RoutingItemId = Convert.ToInt32(newLevel4ID);

                        //MesRouting level1 = GetRoutingByName(caseObj.CASE_C_Data.RoutingName);
                        caseObj.Step4.Level4RoutingItemId = level4RoutingItemId;
                        caseObj.Step4.RoutingItemId = level4RoutingItemId;


                        // 不生成右邊[途程製程], 直接使用現有的
                        //foreach (var x in caseObj.CASE_B_Data.Python_Ret.Data.RountingProcess)
                        foreach (var x in caseObj.Step4.TemplateRoutingItemProcess)
                        {


                            if (true)
                            {

                                var sql5 = $@"

                                INSERT INTO [MES].[RoutingItemProcess]
                                           ([RoutingItemId]
                                           ,[RoutingProcessId]
                                           ,[RoutingItemProcessDesc]
                                           ,[CycleTime]
                                           ,[MoveTime]
                                           ,[Remark]
                                           ,[DisplayStatus]
                                         --  ,[AttrSetting]
                                           ,[CreateDate]
                                           ,[LastModifiedDate]
                                           ,[CreateBy]
                                           ,[LastModifiedBy])
                                     VALUES
                                           ({ caseObj.Step4.RoutingItemId}
                                           ,{ x.RoutingProcessId}
                                           ,'{x.RoutingItemProcessDesc}'
                                           ,{x.CycleTime}
                                           ,{x.MoveTime}
                                           ,'{x.Remark}'
                                           ,'Y' --<DisplayStatus, nvarchar(2),>
                                           --,<AttrSetting, nvarchar(200),>
                                           ,'{strNow}'
                                           ,'{strNow}'
                                           ,{intUser}
                                           ,{intUser}        
                                            )
                                            ; SELECT SCOPE_IDENTITY() AS newLevel5ID
                                        ";
                                caseObj.Step4.DebugSQL5 = sql5.Replace("\r\n", " ");

                                command.CommandText = sql5;
                                object newLevel5ID = command.ExecuteScalar();
                                if (newLevel5ID == null)
                                {
                                    throw new Exception("第5層沒有 能夠寫入");
                                }
                                int level5RoutingItemProcessId = Convert.ToInt32(newLevel5ID);
                                caseObj.Step4.Level5RoutingItemProcessId = level5RoutingItemProcessId;
                            }
                        }

                        //command.CommandText = "INSERT INTO Table2 (Column2) VALUES ('Value2')";
                        //command.ExecuteNonQuery();

                        // Commit the transaction.
                        sqlTransaction.Commit();

                        caseObj.Step4.TransactionEnd = DateTime.Now;
                        caseObj.Step4.TransactionResult = "Transaction committed.";

                        Console.WriteLine("Transaction committed.");
                    }

                    catch (Exception ex)
                    {
                        // Attempt to roll back the transaction.
                        try
                        {
                            sqlTransaction.Rollback();
                            caseObj.Step4.TransactionEnd = DateTime.Now;
                            caseObj.Step4.TransactionResult = $"{ex.Message}  Transaction Rollback.";
                        }
                        catch (Exception exRollback)
                        {
                            // Handle any errors that may have occurred on the server when rolling back the transaction.
                            Console.WriteLine($"Rollback Exception Type: {exRollback.GetType()}");
                            Console.WriteLine($"Rollback Exception: {exRollback.Message}");
                            caseObj.Step4.TransactionResult = "Transaction Rollback. exception  {exRollback.Message}";
                        }

                        Console.WriteLine($"Exception Type: {ex.GetType()}");
                        Console.WriteLine($"Exception: {ex.Message}");
                    }
                }
            }

            // SOP to return 
            // NOTE: not able to use by ref 
            // 無法呼叫控制器 'Business_Manager.Controllers.RoutingController' 上的動作方法 'Void
            // ExecuteTransaction_CASE_D_01(RdDesignControlResult ByRef)'，因為參數 'RdDesignControlResult & amp; caseObj'
            // 是以傳址方式傳遞。
            return caseObj;
        }
        #endregion


        #region //AddRouting-- 新增途程資料 -- Ann 2022-07-14
        public string AddRouting(string RoutingName, string RoutingType, int ModeId)
        {
            try
            {
                if (RoutingName.Length <= 0) throw new SystemException("【途程名稱】不能為空!");
                if (RoutingName.Length > 100) throw new SystemException("【途程名稱】長度錯誤!");
                if (!Regex.IsMatch(RoutingType, "^(1|2)$", RegexOptions.IgnoreCase)) throw new SystemException("【途程類型】錯誤!");
                if (ModeId <= 0) throw new SystemException("【生產模式】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷生產模式是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdMode
                                WHERE ModeId = @ModeId";
                        dynamicParameters.Add("ModeId", ModeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產模式資料錯誤!");
                        #endregion

                        #region //判斷 生產模式+途程名稱 是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE ModeId = @ModeId
                                AND RoutingName = @RoutingName";
                        dynamicParameters.Add("ModeId", ModeId);
                        dynamicParameters.Add("RoutingName", RoutingName);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【生產模式 + 途程名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.Routing (CompanyId, ModeId, RoutingType, RoutingName, Status, RoutingConfirm
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RoutingId
                                VALUES (@CompanyId, @ModeId, @RoutingType, @RoutingName, @Status, @RoutingConfirm
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ModeId,
                                RoutingType,
                                RoutingName,
                                Status = "A",
                                RoutingConfirm = "N",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

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

        #region //AddRoutingProcess-- 新增途程製程資料 -- Ann 2022-07-14
        public string AddRoutingProcess(int RoutingId, string ProcessData)
        {
            try
            {
                if (ProcessData.Length <= 0) throw new SystemException("【製程別名】不能為空!");

                bool checkbool = false;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷途程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ModeId
                                FROM MES.Routing
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程資料錯誤!");

                        int ModeId = -1;
                        foreach (var item in result)
                        {
                            ModeId = item.ModeId;
                        }
                        #endregion

                        #region //判斷途程是否已被確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE RoutingConfirm = @RoutingConfirm
                                AND RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingConfirm", "Y");
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("此途程已被確認，無法更動製程資料!");
                        #endregion

                        var ProcessJson = JObject.Parse(ProcessData);
                        int rowsAffected = 0;
                        foreach (var item in ProcessJson["processInfo"])
                        {
                            #region //判斷製程資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ProcessCheckStatus, a.ProcessCheckType, a.PackageFlag, a.ConsumeFlag
                                    , b.ProcessName
                                    FROM MES.ProcessParameter a 
                                    INNER JOIN MES.Process b ON a.ProcessId = b.ProcessId
                                    WHERE a.ProcessId = @ProcessId
                                    AND a.ModeId = @ModeId";
                            dynamicParameters.Add("ProcessId", item["ProcessId"].ToString());
                            dynamicParameters.Add("ModeId", ModeId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("製程資料錯誤!");

                            string ProcessCheckStatus = "";
                            string ProcessCheckType = "";
                            string PackageFlag = "";
                            string ProcessName = "";
                            string ConsumeFlag = "";

                            foreach (var item2 in result3)
                            {
                                if (item2.ProcessCheckStatus == null) throw new SystemException("製程【" + item2.ProcessName + "】未設定製程參數!");
                                ProcessCheckStatus = item2.ProcessCheckStatus;
                                ProcessCheckType = item2.ProcessCheckType;
                                PackageFlag = item2.PackageFlag;
                                ProcessName = item2.ProcessName;
                                ConsumeFlag = item2.ConsumeFlag;

                            }
                            #endregion

                            if (Convert.ToInt32(item["ProcessId"]) == 14) checkbool = true;

                            #region //取得最大排序數
                            int maxSortNumber = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSortNumber
                                    FROM MES.RoutingProcess
                                    WHERE RoutingId = @RoutingId";
                            dynamicParameters.Add("RoutingId", RoutingId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item2 in result4)
                            {
                                maxSortNumber = Convert.ToInt32(item2.MaxSortNumber) + 1;
                            }
                            #endregion

                            #region //批次新增單一站多個
                            for (int i = 0; i < Convert.ToInt32(item["CreateCount"]); i++)
                            {
                                string ProcessAlias = "";

                                #region //確認是否有相同製程，若有則自動帶【-】
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT COUNT(a.RoutingProcessId) Count
                                            FROM MES.RoutingProcess a
                                            WHERE a.RoutingId = @RoutingId
                                            AND a.ProcessId = @ProcessId";
                                dynamicParameters.Add("RoutingId", RoutingId);
                                dynamicParameters.Add("ProcessId", item["ProcessId"].ToString());

                                var RoutingProcessResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item2 in RoutingProcessResult)
                                {
                                    if (item2.Count > 0)
                                    {
                                        ProcessAlias = item["ProcessName"].ToString() + "-" + (item2.Count + 1).ToString();
                                    }
                                    else
                                    {
                                        ProcessAlias = item["ProcessName"].ToString();
                                    }
                                }
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.RoutingProcess (RoutingId, ProcessId, SortNumber, ProcessAlias, DisplayStatus, NecessityStatus
                                        , ProcessCheckStatus, ProcessCheckType, Status, PackageFlag , ConsumeFlag
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.RoutingProcessId
                                        VALUES (@RoutingId, @ProcessId, @SortNumber, @ProcessAlias, @DisplayStatus, @NecessityStatus
                                        , @ProcessCheckStatus, @ProcessCheckType, @Status, @PackageFlag, @ConsumeFlag
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RoutingId,
                                        ProcessId = Convert.ToInt32(item["ProcessId"]),
                                        SortNumber = maxSortNumber + i,
                                        ProcessAlias,
                                        DisplayStatus = "Y",
                                        NecessityStatus = "Y",
                                        ProcessCheckStatus,
                                        ProcessCheckType,
                                        Status = "A",
                                        PackageFlag,
                                        ConsumeFlag,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                int RoutingProcessId = -1;
                                foreach (var item3 in insertResult)
                                {
                                    RoutingProcessId = item3.RoutingProcessId;
                                }

                                #region //若為刻字站，自動增加刻字屬性
                                if (ProcessName == "刻字")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.RoutingProcessItem (RoutingProcessId, ItemNo, ChkUnique
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.RoutingProcessItemId
                                            VALUES (@RoutingProcessId, @ItemNo, @ChkUnique
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            RoutingProcessId,
                                            ItemNo = "Lettering",
                                            ChkUnique = "Y",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                                #endregion
                            }
                            #endregion
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

                if (checkbool == true)
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //取得目前製程順序
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT RoutingProcessId, SortNumber
                                    FROM MES.RoutingProcess
                                    WHERE RoutingId = @RoutingId
                                    ORDER BY SortNumber";
                            dynamicParameters.Add("RoutingId", RoutingId);

                            var result = sqlConnection.Query(sql, dynamicParameters);

                            #region //先將全部製程順序改為-1
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.RoutingProcess SET
                                    SortNumber = SortNumber * -1,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE RoutingId = @RoutingId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    RoutingId
                                });

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            foreach (var item in result)
                            {
                                #region //先將製程順序全部+1
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.RoutingProcess SET
                                        SortNumber = @SortNumber,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE RoutingProcessId = @RoutingProcessId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SortNumber = item.SortNumber + 1,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        item.RoutingProcessId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion

                            #region //新增退鎳站
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.RoutingProcess (RoutingId, ProcessId, SortNumber, ProcessAlias, DisplayStatus, NecessityStatus
                                    , ProcessCheckStatus, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@RoutingId, @ProcessId, @SortNumber, @ProcessAlias, @DisplayStatus, @NecessityStatus
                                    , @ProcessCheckStatus, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RoutingId,
                                    ProcessId = 15,
                                    SortNumber = 1,
                                    ProcessAlias = "退鎳",
                                    DisplayStatus = "N",
                                    NecessityStatus = "N",
                                    ProcessCheckStatus = "N",
                                    Status = "A",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion
                        }
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

        #region //AddBatchRoutingProcess-- 批量新增途程製程資料 -- Ann 2024-09-09
        public string AddBatchRoutingProcess(int RoutingId, string UploadJson)
        {
            try
            {
                if (UploadJson.Length <= 0) throw new SystemException("【途程資料】不能為空!");
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷途程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ModeId, a.RoutingConfirm
                                FROM MES.Routing a
                                WHERE a.RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var RoutingResult = sqlConnection.Query(sql, dynamicParameters);
                        if (RoutingResult.Count() <= 0) throw new SystemException("途程資料錯誤!");

                        int ModeId = -1;
                        foreach (var item in RoutingResult)
                        {
                            if (item.RoutingConfirm != "N") throw new SystemException("此途程已確認，無法更改!!");
                            ModeId = item.ModeId;
                        }
                        #endregion

                        #region //確認原本是否已有製程
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.RoutingProcess a 
                                WHERE a.RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var RoutingProcessResult = sqlConnection.Query(sql, dynamicParameters);

                        if (RoutingProcessResult.Count() > 0) throw new SystemException("此途程已建立過製程，無法使用批量新增的方式!!");
                        #endregion

                        JObject routingProcessJson = JObject.Parse(UploadJson);

                        int count = 1;
                        foreach (var item in routingProcessJson["routingProcessInfo"])
                        {
                            #region //資料確認
                            string ProcessNo = "";
                            Regex regex = new Regex("^[YN]$");
                            if (item["ProcessNo"] != null)
                            {
                                ProcessNo = item["ProcessNo"].ToString();
                            }
                            else
                            {
                                throw new SystemException("製程編號不能為空!!");
                            }

                            string ProcessAlias = "";
                            if (item["ProcessAlias"] != null)
                            {
                                ProcessAlias = item["ProcessAlias"].ToString();
                            }
                            else
                            {
                                throw new SystemException("製程別名不能為空!!");
                            }

                            int SortNumber = -1;
                            if (item["SortNumber"] != null)
                            {
                                if (int.TryParse(item["SortNumber"].ToString(), out _))
                                {
                                    SortNumber = Convert.ToInt32(item["SortNumber"]);
                                }
                                else
                                {
                                    throw new SystemException("製程順序必須為數字格式!!");
                                }
                            }
                            else
                            {
                                throw new SystemException("製程順序不能為空!!");
                            }

                            string ProcessCheckStatus = "";
                            if (item["ProcessCheckStatus"] != null)
                            {
                                if (regex.IsMatch(item["ProcessCheckStatus"].ToString()))
                                {
                                    ProcessCheckStatus = item["ProcessCheckStatus"].ToString();
                                }
                                else
                                {
                                    throw new SystemException("是否支援工程檢格式錯誤!!");
                                }
                            }
                            else
                            {
                                throw new SystemException("是否支援工程檢不能為空!!");
                            }

                            string ProcessCheckType = "";
                            if (item["ProcessCheckStatus"].ToString() == "Y" && item["ProcessCheckType"] == null)
                            {
                                throw new SystemException("若需工程檢，則工程檢頻率不能為空!!");
                            }

                            if (item["ProcessCheckType"] != null)
                            {
                                if (item["ProcessCheckType"].Count() > 0)
                                {
                                    if (item["ProcessCheckType"].ToString() != "1" && item["ProcessCheckType"].ToString() != "2")
                                    {
                                        throw new SystemException("工程檢頻率格式錯誤!!");
                                    }

                                    ProcessCheckType = item["ProcessCheckType"].ToString();
                                }
                            }

                            string DisplayStatus = "";
                            if (item["DisplayStatus"] != null)
                            {
                                if (regex.IsMatch(item["DisplayStatus"].ToString()))
                                {
                                    DisplayStatus = item["DisplayStatus"].ToString();
                                }
                                else
                                {
                                    throw new SystemException("是否顯示在流程卡格式錯誤!!");
                                }
                            }
                            else
                            {
                                throw new SystemException("是否顯示在流程卡不能為空!!");
                            }

                            string NecessityStatus = "";
                            if (item["NecessityStatus"] != null)
                            {
                                if (regex.IsMatch(item["NecessityStatus"].ToString()))
                                {
                                    NecessityStatus = item["NecessityStatus"].ToString();
                                }
                                else
                                {
                                    throw new SystemException("是否為必要過站格式錯誤!!");
                                }
                            }
                            else
                            {
                                throw new SystemException("是否為必要過站不能為空!!");
                            }
                            #endregion

                            #region //判斷製程資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ProcessId, a.ProcessCheckStatus, a.ProcessCheckType, a.PackageFlag
                                    , b.ProcessName
                                    FROM MES.ProcessParameter a 
                                    INNER JOIN MES.Process b ON a.ProcessId = b.ProcessId
                                    WHERE a.ModeId = @ModeId
                                    AND b.ProcessNo = @ProcessNo";
                            dynamicParameters.Add("ModeId", ModeId);
                            dynamicParameters.Add("ProcessNo", ProcessNo);

                            var ProcessParameterResult = sqlConnection.Query(sql, dynamicParameters);
                            if (ProcessParameterResult.Count() <= 0) throw new SystemException("製程資料錯誤!");

                            int ProcessId = -1;
                            string ProcessName = "";
                            foreach (var item2 in ProcessParameterResult)
                            {
                                if (item2.ProcessCheckStatus == null) throw new SystemException("製程【" + item2.ProcessName + "】未設定製程參數!");
                                ProcessId = item2.ProcessId;
                                ProcessName = item2.ProcessName;
                            }
                            #endregion

                            #region //確認號碼有無連續
                            if (count != SortNumber)
                            {
                                throw new SystemException("製程順序沒有連續，請重新維護!!");
                            }
                            #endregion

                            #region //確認是否有相同製程，若有則自動帶【-】
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT COUNT(a.RoutingProcessId) Count
                                    FROM MES.RoutingProcess a
                                    WHERE a.RoutingId = @RoutingId
                                    AND a.ProcessId = @ProcessId";
                            dynamicParameters.Add("RoutingId", RoutingId);
                            dynamicParameters.Add("ProcessId", ProcessId);

                            var CheckRoutingProcessResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item2 in CheckRoutingProcessResult)
                            {
                                if (item2.Count > 0)
                                {
                                    ProcessAlias = ProcessName + "-" + (item2.Count + 1).ToString();
                                }
                                else
                                {
                                    ProcessAlias = ProcessName.ToString();
                                }
                            }
                            #endregion

                            #region //Insert MES.RoutingProcess
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.RoutingProcess (RoutingId, ProcessId, SortNumber, ProcessAlias, DisplayStatus, NecessityStatus
                                    , ProcessCheckStatus, ProcessCheckType, Status, PackageFlag
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.RoutingProcessId
                                    VALUES (@RoutingId, @ProcessId, @SortNumber, @ProcessAlias, @DisplayStatus, @NecessityStatus
                                    , @ProcessCheckStatus, @ProcessCheckType, @Status, @PackageFlag
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RoutingId,
                                    ProcessId,
                                    SortNumber,
                                    ProcessAlias,
                                    DisplayStatus,
                                    NecessityStatus,
                                    ProcessCheckStatus,
                                    ProcessCheckType,
                                    Status = "A",
                                    PackageFlag = "N",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            int RoutingProcessId = -1;
                            foreach (var item3 in insertResult)
                            {
                                RoutingProcessId = item3.RoutingProcessId;
                            }

                            #region //若為刻字站，自動增加刻字屬性
                            if (ProcessName == "刻字")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.RoutingProcessItem (RoutingProcessId, ItemNo, ChkUnique
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.RoutingProcessItemId
                                        VALUES (@RoutingProcessId, @ItemNo, @ChkUnique
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RoutingProcessId,
                                        ItemNo = "Lettering",
                                        ChkUnique = "Y",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
                            #endregion
                            #endregion

                            count++;
                        }

                        #region //統整目前途程製程資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RoutingProcessId, a.RoutingId, a.ProcessId, a.SortNumber, a.ProcessAlias
                                , a.DisplayStatus, a.NecessityStatus, a.ProcessCheckStatus, a.ProcessCheckType, a.PackageFlag
                                , b.ProcessNo, b.ProcessName
                                FROM MES.RoutingProcess a 
                                INNER JOIN MES.Process b ON a.ProcessId = b.ProcessId
                                WHERE a.RoutingId = @RoutingId
                                ORDER BY SortNumber";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var Result = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = Result
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

        #region //AddRoutingItem-- 新增途程品號資料 -- Ann 2022-07-19
        public string AddRoutingItem(int RoutingId, int ControlId, int MtlItemId, string Status)
        {
            try
            {
                if (!Regex.IsMatch(Status, "^(A|S)$", RegexOptions.IgnoreCase)) throw new SystemException("【狀態】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷途程資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.Routing
                                    WHERE RoutingId = @RoutingId";
                            dynamicParameters.Add("RoutingId", RoutingId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("途程資料錯誤!");
                            #endregion

                            #region //判斷研發設計圖版控資料是否正確
                            if (ControlId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM PDM.RdDesignControl
                                        WHERE ControlId = @ControlId";
                                dynamicParameters.Add("ControlId", ControlId);

                                var result2 = sqlConnection.Query(sql, dynamicParameters);
                                if (result2.Count() <= 0) throw new SystemException("研發設計圖版控資料錯誤!");
                            }
                            #endregion

                            #region //判斷品號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemNo, TransferStatus
                                    FROM PDM.MtlItem
                                    WHERE MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("品號資料錯誤!");

                            string MtlItemNo = "";
                            string TransferStatus = "";
                            foreach (var item in result3)
                            {
                                if (item.TransferStatus != "Y") throw new SystemException("此品號尚未拋轉ERP!!");
                                MtlItemNo = item.MtlItemNo;
                                TransferStatus = item.TransferStatus;
                            }
                            #endregion

                            #region //判斷ERP品號生效日與失效日
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB030)) MB030, LTRIM(RTRIM(MB031)) MB031
                                    FROM INVMB
                                    WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVMBResult.Count() <= 0) throw new SystemException("此品號不存在於ERP中!!");

                            foreach (var item in INVMBResult)
                            {
                                if (item.MB030 != "" && item.MB030 != null)
                                {
                                    #region //判斷日期需大於或等於生效日
                                    string EffectiveDate = item.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
                                    if (effresult < 0) throw new SystemException("此品號尚未生效，無法使用!!");
                                    #endregion
                                }

                                if (item.MB031 != "" && item.MB031 != null)
                                {
                                    #region //判斷日期需小於或等於失效日
                                    string ExpirationDate = item.MB031;
                                    string effYear = ExpirationDate.Substring(0, 4);
                                    string effMonth = ExpirationDate.Substring(4, 2);
                                    string effDay = ExpirationDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                    #endregion
                                }
                            }
                            #endregion

                            #region //判斷 途程+研發設計圖版控 是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.RoutingItem
                                    WHERE RoutingId = @RoutingId
                                    AND ControlId = @ControlId";
                            dynamicParameters.Add("RoutingId", RoutingId);
                            dynamicParameters.Add("ControlId", ControlId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);
                            if (result4.Count() > 0) throw new SystemException("【途程 + 研發設計圖版控】重複，請重新輸入!");
                            #endregion

                            #region //INSERT MES.RoutingItem
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.RoutingItem (RoutingId, ControlId, MtlItemId, Status, RoutingItemConfirm
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.RoutingItemId
                                    VALUES (@RoutingId, @ControlId, @MtlItemId, @Status, @RoutingItemConfirm
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RoutingId,
                                    ControlId = ControlId > 0 ? (int?)ControlId : null,
                                    MtlItemId,
                                    Status,
                                    RoutingItemConfirm = "N",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            int RoutingItemId = 0;
                            foreach (var item in insertResult)
                            {
                                RoutingItemId = Convert.ToInt32(item.RoutingItemId);
                            }
                            #endregion

                            #region //塞加工屬性
                            #region //找此張設計圖加工屬性
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.AttributeId, a.AttributeItemId, a.AttributeValue
                                    , a.UpperLimit, a.LowerLimit
                                    FROM PDM.RdDesignAttribute a
                                    WHERE a.ControlId = @ControlId";
                            dynamicParameters.Add("ControlId", ControlId);

                            var result5 = sqlConnection.Query(sql, dynamicParameters);
                            if (result5.Count() > 0)
                            {
                                foreach (var item in result5)
                                {
                                    //開始塞資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.RoutingItemAttribute (RoutingItemId, AttributeItemId
                                            , AttributeValue, UpperLimit, LowerLimit
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES (@RoutingItemId, @AttributeItemId
                                            , @AttributeValue, @UpperLimit, @LowerLimit
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            RoutingItemId,
                                            item.AttributeItemId,
                                            item.AttributeValue,
                                            item.UpperLimit,
                                            item.LowerLimit,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion
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

        #region //AddBatchRoutingItem-- 批量新增途程品號資料 -- Ann 2024-09-10
        public string AddBatchRoutingItem(int RoutingId, string UploadJson)
        {
            try
            {
                if (UploadJson.Length <= 0) throw new SystemException("途程品號資料不能為空!!");
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷途程資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.RoutingConfirm 
                                    , b.ModeNo
                                    FROM MES.Routing a 
                                    INNER JOIN MES.ProdMode b ON a.ModeId = b.ModeId
                                    WHERE a.RoutingId = @RoutingId";
                            dynamicParameters.Add("RoutingId", RoutingId);

                            var RoutingResult = sqlConnection.Query(sql, dynamicParameters);
                            if (RoutingResult.Count() <= 0) throw new SystemException("途程資料錯誤!");

                            string ModeNo = "";
                            foreach (var item in RoutingResult)
                            {
                                if (item.RoutingConfirm != "Y") throw new SystemException("途程尚未確認，無法新增!!");
                                ModeNo = item.ModeNo;
                            }
                            #endregion

                            JObject routingItemJson = JObject.Parse(UploadJson);
                            List<int> MtlItemIdList = new List<int>();
                            List<int> RoutingProcessIdList = new List<int>();
                            bool mtlItemBool = false; //是否為新品號，預設為否
                            int RoutingItemId = 0;
                            Regex regex = new Regex(@"^-?\d+(\.\d+)?$");
                            int count = 1;
                            foreach (var item in routingItemJson["routingItemInfo"])
                            {
                                #region //資料驗證
                                #region //品號
                                mtlItemBool = false;
                                int MtlItemId = -1;
                                string MtlItemNo = item["MtlItemNo"].ToString();
                                string TransferStatus = "";
                                if (item["MtlItemNo"] != null)
                                {
                                    #region //判斷品號資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT MtlItemId, MtlItemNo, TransferStatus
                                            FROM PDM.MtlItem
                                            WHERE MtlItemNo = @MtlItemNo";
                                    dynamicParameters.Add("MtlItemNo", MtlItemNo);

                                    var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);
                                    if (MtlItemResult.Count() <= 0) throw new SystemException("品號【" + MtlItemNo + "】資料錯誤!");

                                    foreach (var item2 in MtlItemResult)
                                    {
                                        if (item2.TransferStatus != "Y") throw new SystemException("品號【" + MtlItemNo + "】尚未拋轉ERP!!");
                                        if (MtlItemIdList.IndexOf(item2.MtlItemId) == -1)
                                        {
                                            MtlItemIdList.Add(item2.MtlItemId);
                                            mtlItemBool = true;
                                        }

                                        MtlItemId = item2.MtlItemId;
                                        TransferStatus = item2.TransferStatus;
                                    }
                                    #endregion

                                    #region //判斷ERP品號生效日與失效日
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT LTRIM(RTRIM(MB030)) MB030, LTRIM(RTRIM(MB031)) MB031
                                            FROM INVMB
                                            WHERE MB001 = @MB001";
                                    dynamicParameters.Add("MB001", MtlItemNo);

                                    var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (INVMBResult.Count() <= 0) throw new SystemException("品號【" + MtlItemNo + "】不存在於ERP中!!");

                                    foreach (var item2 in INVMBResult)
                                    {
                                        if (item2.MB030 != "" && item2.MB030 != null)
                                        {
                                            #region //判斷日期需大於或等於生效日
                                            string EffectiveDate = item2.MB030;
                                            string effYear = EffectiveDate.Substring(0, 4);
                                            string effMonth = EffectiveDate.Substring(4, 2);
                                            string effDay = EffectiveDate.Substring(6, 2);
                                            DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                            int effresult = DateTime.Compare(CreateDate, effFullDate);
                                            if (effresult < 0) throw new SystemException("品號【" + MtlItemNo + "】尚未生效，無法使用!!");
                                            #endregion
                                        }

                                        if (item2.MB031 != "" && item2.MB031 != null)
                                        {
                                            #region //判斷日期需小於或等於失效日
                                            string ExpirationDate = item2.MB031;
                                            string effYear = ExpirationDate.Substring(0, 4);
                                            string effMonth = ExpirationDate.Substring(4, 2);
                                            string effDay = ExpirationDate.Substring(6, 2);
                                            DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                            int effresult = DateTime.Compare(CreateDate, effFullDate);
                                            if (effresult > 0) throw new SystemException("品號【" + MtlItemNo + "】已失效，無法使用!!");
                                            #endregion
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    throw new SystemException("品號不能為空!!");
                                }
                                #endregion

                                #region //設計圖
                                int ControlId = -1;
                                if (item["RdDesign"] != null)
                                {
                                    string rdDesignType = item["RdDesign"].Type.ToString();
                                    if (rdDesignType != "Null")
                                    {
                                        string rdDesign = item["RdDesign"].ToString();
                                        if (rdDesign.IndexOf(":") == -1)
                                        {
                                            throw new SystemException("品號【" + MtlItemNo + "】設計圖格式錯誤!!");
                                        }
                                        else
                                        {
                                            ControlId = Convert.ToInt32(rdDesign.Split(':')[0]);
                                        }
                                    }
                                    else
                                    {
                                        if (ModeNo == "JMO-A-001")
                                        {
                                            throw new SystemException("品號【" + MtlItemNo + "】設計圖不能為空!!");
                                        }
                                    }
                                }
                                else
                                {
                                    if (ModeNo == "JMO-A-001")
                                    {
                                        throw new SystemException("品號【" + MtlItemNo + "】設計圖不能為空!!");
                                    }
                                }

                                #region //確認此設計途版控資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ReleasedStatus
                                        FROM PDM.RdDesignControl a 
                                        WHERE a.ControlId = @ControlId";
                                dynamicParameters.Add("ControlId", ControlId);

                                var RdDesignControlResult = sqlConnection.Query(sql, dynamicParameters);

                                if (RdDesignControlResult.Count() <= 0) throw new SystemException("品號【" + MtlItemNo + "】設計圖資料有誤!!");

                                foreach (var item2 in RdDesignControlResult)
                                {
                                    if (item2.ReleasedStatus != "Y")
                                    {
                                        throw new SystemException("品號【" + MtlItemNo + "】設計圖目前非發行狀態!!");
                                    }
                                }
                                #endregion

                                #region //判斷 途程+研發設計圖版控 是否重複
                                if (mtlItemBool == true)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MES.RoutingItem
                                            WHERE RoutingId = @RoutingId
                                            AND ControlId = @ControlId";
                                    dynamicParameters.Add("RoutingId", RoutingId);
                                    dynamicParameters.Add("ControlId", ControlId);

                                    var RoutingItemResult = sqlConnection.Query(sql, dynamicParameters);
                                    if (RoutingItemResult.Count() > 0) throw new SystemException("品號【" + MtlItemNo + "】<br>【途程 + 研發設計圖版控】重複，請重新輸入!");
                                }
                                #endregion
                                #endregion

                                #region //製程
                                int RoutingProcessId = -1;
                                if (item["RoutingProcess"] != null)
                                {
                                    string routingProcess = item["RoutingProcess"].ToString();
                                    if (routingProcess.IndexOf(":") == -1)
                                    {
                                        throw new SystemException("品號【" + MtlItemNo + "】<br>製程【" + routingProcess + "】格式錯誤!!");
                                    }
                                    else
                                    {
                                        RoutingProcessId = Convert.ToInt32(routingProcess.Split(':')[0]);
                                    }

                                    #region //確認途程製程資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MES.RoutingProcess a 
                                            WHERE a.RoutingProcessId = @RoutingProcessId";
                                    dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                                    var RoutingProcessResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (RoutingProcessResult.Count() <= 0) throw new SystemException("品號【" + MtlItemNo + "】<br>製程【" + routingProcess + "】資料錯誤!!");

                                    if (RoutingProcessIdList.IndexOf(RoutingProcessId) == -1)
                                    {
                                        RoutingProcessIdList.Add(RoutingProcessId);
                                    }
                                    #endregion
                                }
                                else
                                {
                                    throw new SystemException("品號【" + MtlItemNo + "】製程不能為空!!");
                                }
                                #endregion

                                #region //製程說明
                                string RoutingItemProcessDesc = "";
                                if (item["RoutingItemProcessDesc"] != null)
                                {
                                    RoutingItemProcessDesc = item["RoutingItemProcessDesc"].ToString();
                                }
                                #endregion

                                #region //標準工時
                                int CycleTime = 0;
                                if (item["CycleTime"] != null)
                                {
                                    if (regex.IsMatch(item["CycleTime"].ToString()))
                                    {
                                        CycleTime = Convert.ToInt32(item["CycleTime"]);
                                    }
                                    else
                                    {
                                        throw new SystemException("品號【" + MtlItemNo + "】完工時間格式錯誤!");
                                    }
                                }
                                #endregion

                                #region //上下製程移轉時間
                                int MoveTime = 0;
                                if (item["MoveTime"] != null)
                                {
                                    if (regex.IsMatch(item["MoveTime"].ToString()))
                                    {
                                        MoveTime = Convert.ToInt32(item["MoveTime"]);
                                    }
                                    else
                                    {
                                        throw new SystemException("品號【" + MtlItemNo + "】上下製程移轉時間格式錯誤!");
                                    }
                                }
                                #endregion

                                #region //標準工時上下公差
                                int ProcessingTime = 0;
                                if (item["ProcessingTime"] != null)
                                {
                                    if (regex.IsMatch(item["ProcessingTime"].ToString()))
                                    {
                                        ProcessingTime = Convert.ToInt32(item["ProcessingTime"]);
                                    }
                                    else
                                    {
                                        throw new SystemException("品號【" + MtlItemNo + "】標準工時上下公差格式錯誤!");
                                    }
                                }
                                #endregion

                                #region //當站最短完工時間
                                int WatingTime = 0;
                                if (item["WatingTime"] != null)
                                {
                                    if (regex.IsMatch(item["WatingTime"].ToString()))
                                    {
                                        WatingTime = Convert.ToInt32(item["WatingTime"]);
                                    }
                                    else
                                    {
                                        throw new SystemException("品號【" + MtlItemNo + "】當站最短完工時間格式錯誤!");
                                    }
                                }
                                #endregion

                                #region //加工備註
                                string Remark = "";
                                if (item["Remark"] != null)
                                {
                                    Remark = item["Remark"].ToString();
                                }
                                #endregion

                                #region //是否顯示於流程卡
                                string DisplayStatus = "";
                                if (item["DisplayStatus"] != null)
                                {
                                    DisplayStatus = item["DisplayStatus"].ToString();
                                }
                                #endregion
                                #endregion

                                #region //建立途程品號
                                if (mtlItemBool == true)
                                {
                                    #region //Insert MES.RoutingItem
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.RoutingItem (RoutingId, ControlId, MtlItemId, Status, RoutingItemConfirm
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.RoutingItemId
                                            VALUES (@RoutingId, @ControlId, @MtlItemId, @Status, @RoutingItemConfirm
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            RoutingId,
                                            ControlId = ControlId > 0 ? (int?)ControlId : null,
                                            MtlItemId,
                                            Status = "A",
                                            RoutingItemConfirm = "Y",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();

                                    foreach (var item2 in insertResult)
                                    {
                                        RoutingItemId = Convert.ToInt32(item2.RoutingItemId);
                                    }
                                    #endregion

                                    #region //塞加工屬性
                                    #region //找此張設計圖加工屬性
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.AttributeId, a.AttributeItemId, a.AttributeValue
                                            , a.UpperLimit, a.LowerLimit
                                            FROM PDM.RdDesignAttribute a
                                            WHERE a.ControlId = @ControlId";
                                    dynamicParameters.Add("ControlId", ControlId);

                                    var RdDesignAttributeResult = sqlConnection.Query(sql, dynamicParameters);

                                    foreach (var item2 in RdDesignAttributeResult)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO MES.RoutingItemAttribute (RoutingItemId, AttributeItemId
                                                , AttributeValue, UpperLimit, LowerLimit
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                VALUES (@RoutingItemId, @AttributeItemId
                                                , @AttributeValue, @UpperLimit, @LowerLimit
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                RoutingItemId,
                                                item2.AttributeItemId,
                                                item2.AttributeValue,
                                                item2.UpperLimit,
                                                item2.LowerLimit,
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });

                                        insertResult = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult.Count();
                                    }
                                    #endregion
                                    #endregion

                                    #region //檢核製程數量是否與途程製程相同
                                    if (count > 1)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT COUNT(1) Count
                                                FROM MES.RoutingProcess a 
                                                WHERE a.RoutingId = @RoutingId";
                                        dynamicParameters.Add("RoutingId", RoutingId);

                                        var RoutingProcessResult = sqlConnection.Query(sql, dynamicParameters);

                                        foreach (var item2 in RoutingProcessResult)
                                        {
                                            if (item2.Count != RoutingProcessIdList.Count())
                                            {
                                                throw new SystemException("製程數量異常，請重新帶入資料!!");
                                            }
                                        }
                                    }
                                    #endregion
                                }
                                #endregion

                                #region //建立加工流程卡資料
                                #region //確認此途程品號無重複製程
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM MES.RoutingItemProcess a 
                                        WHERE a.RoutingItemId = @RoutingItemId 
                                        AND a.RoutingProcessId = @RoutingProcessId";
                                dynamicParameters.Add("RoutingItemId", RoutingItemId);
                                dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                                var RoutingItemProcessResult = sqlConnection.Query(sql, dynamicParameters);

                                if (RoutingItemProcessResult.Count() > 0) throw new SystemException("品號【" + MtlItemNo + "】重複建立流程卡製程，請確認是否有重複選取品號!");
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.RoutingItemProcess (RoutingItemId, RoutingProcessId, RoutingItemProcessDesc, CycleTime, MoveTime
                                        , ProcessingTime, WatingTime, Remark, DisplayStatus
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.RoutingItemId
                                        VALUES (@RoutingItemId, @RoutingProcessId, @RoutingItemProcessDesc, @CycleTime, @MoveTime
                                        , @ProcessingTime, @WatingTime, @Remark, @DisplayStatus
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RoutingItemId,
                                        RoutingProcessId,
                                        RoutingItemProcessDesc,
                                        CycleTime,
                                        MoveTime,
                                        ProcessingTime,
                                        WatingTime,
                                        Remark,
                                        DisplayStatus,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult2.Count();
                                #endregion

                                count++;
                            }

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

        #region //AddCopyRoutingItem-- 複製途程品號資料 -- Ann 2022-07-19
        public string AddCopyRoutingItem(int RoutingItemId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷途程品號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RoutingId, a.ControlId, a.MtlItemId, a.[Status], a.RoutingItemConfirm
                                FROM MES.RoutingItem a
                                WHERE a.RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程品號資料錯誤!");

                        int RoutingId = 0;
                        int MtlItemId = 0;
                        foreach (var item in result)
                        {
                            RoutingId = item.RoutingId;
                            MtlItemId = item.MtlItemId;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.RoutingItem (RoutingId, MtlItemId, Status, RoutingItemConfirm
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RoutingItemId
                                VALUES (@RoutingId, @MtlItemId, @Status, @RoutingItemConfirm
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoutingId,
                                MtlItemId,
                                Status = "A",
                                RoutingItemConfirm = "N",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var data = sqlConnection.Query(sql, dynamicParameters);

                        int newRoutingItemId = -1;
                        foreach (var item in data)
                        {
                            newRoutingItemId = item.RoutingItemId;
                        }

                        int rowsAffected = data.Count();

                        #region //同步新增流程卡細節內容
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RoutingProcessId, a.RoutingItemProcessDesc, a.Remark, a.DisplayStatus
                                FROM MES.RoutingItemProcess a
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in result2)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.RoutingItemProcess (RoutingItemId, RoutingProcessId, RoutingItemProcessDesc, Remark, DisplayStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.RoutingItemId
                                    VALUES (@RoutingItemId, @RoutingProcessId, @RoutingItemProcessDesc, @Remark, @DisplayStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RoutingItemId = newRoutingItemId,
                                    item.RoutingProcessId,
                                    item.RoutingItemProcessDesc,
                                    item.Remark,
                                    item.DisplayStatus,                                    
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

        #region //AddCopyRouting-- 複製途程資料 -- Ann 2022-07-25
        public string AddCopyRouting(int RoutingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷途程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId, a.ModeId, a.RoutingType, a.RoutingName, a.[Status]
                                FROM MES.Routing a
                                WHERE a.RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程資料錯誤!");

                        int CompanyId = 0;
                        int ModeId = 0;
                        string RoutingType = "";
                        string RoutingName = "";
                        string Status = "";
                        foreach (var item in result)
                        {
                            CompanyId = item.CompanyId;
                            ModeId = item.ModeId;
                            RoutingType = item.RoutingType;
                            RoutingName = "Copy From: " + item.RoutingName;
                            Status = item.Status;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.Routing (CompanyId, ModeId, RoutingType, RoutingName, Status, RoutingConfirm
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RoutingId
                                VALUES (@CompanyId, @ModeId, @RoutingType, @RoutingName, @Status, @RoutingConfirm
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId,
                                ModeId,
                                RoutingType,
                                RoutingName,
                                Status,
                                RoutingConfirm = "N",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        int newRoutingId = 0;
                        foreach (var item in insertResult)
                        {
                            newRoutingId = item.RoutingId;
                        }

                        #region //複製製程
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RoutingId, a.ProcessId, a.SortNumber, a.ProcessAlias
                                , a.DisplayStatus, a.NecessityStatus, a.ProcessCheckStatus, a.ProcessCheckType, a.Status
                                , (
                                  SELECT b.ItemNo, b.ChkUnique
                                  FROM MES.RoutingProcessItem b
                                  INNER JOIN MES.RoutingProcess ba ON b.RoutingProcessId = ba.RoutingProcessId
                                  WHERE b.RoutingProcessId = a.RoutingProcessId
                                  FOR JSON PATH, ROOT('data')
                                ) RoutingProcessItem
                                FROM MES.RoutingProcess a
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0)
                        {
                            foreach (var item in result2)
                            {
                                int ProcessId = item.ProcessId;
                                int SortNumber = item.SortNumber;
                                string ProcessAlias = item.ProcessAlias;
                                string DisplayStatus = item.DisplayStatus;
                                string NecessityStatus = item.NecessityStatus;
                                string ProcessCheckStatus = item.ProcessCheckStatus;
                                string ProcessCheckType = item.ProcessCheckType;

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.RoutingProcess (RoutingId, ProcessId, SortNumber, ProcessAlias, DisplayStatus, NecessityStatus, ProcessCheckStatus, ProcessCheckType, Status
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.RoutingProcessId
                                        VALUES (@RoutingId, @ProcessId, @SortNumber, @ProcessAlias, @DisplayStatus, @NecessityStatus, @ProcessCheckStatus, @ProcessCheckType, @Status
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RoutingId = newRoutingId,
                                        ProcessId,
                                        SortNumber,
                                        ProcessAlias,
                                        DisplayStatus,
                                        NecessityStatus,
                                        ProcessCheckStatus,
                                        ProcessCheckType,
                                        Status,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                int RoutingProcessId = 1;
                                foreach (var item2 in insertResult)
                                {
                                    RoutingProcessId = item2.RoutingProcessId;
                                }

                                if (item.RoutingProcessItem != null)
                                {
                                    var routingProcessItemJson = JObject.Parse(item.RoutingProcessItem);

                                    foreach (var item3 in routingProcessItemJson.data)
                                    {
                                        #region //複製途程製程屬性
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO MES.RoutingProcessItem (RoutingProcessId, ItemNo, ChkUnique
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.RoutingProcessId, INSERTED.ItemNo
                                                VALUES (@RoutingProcessId, @ItemNo, @ChkUnique
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                RoutingProcessId,
                                                ItemNo = item3["ItemNo"].ToString(),
                                                ChkUnique = item3["ChkUnique"].ToString(),
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });

                                        insertResult = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult.Count();
                                        #endregion
                                    }
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

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddRoutingProcessItem-- 新增途程製程屬性檔 -- Ann 2022-08-12
        public string AddRoutingProcessItem(int RoutingProcessId, string ItemNo, string ChkUnique)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ChkUnique.Length <= 0) throw new SystemException("【是否檢核重複值】不能為空!");

                        #region //判斷途程製程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 RoutingId
                                FROM MES.RoutingProcess a
                                WHERE a.RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程製程資料錯誤!");

                        int RoutingId = 0;
                        foreach (var item in result)
                        {
                            RoutingId = item.RoutingId;
                        }
                        #endregion

                        #region //判斷類別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Type a
                                WHERE a.TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeNo", ItemNo);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("類別資料錯誤!");
                        #endregion

                        #region //判斷 途程製程+項目 是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.RoutingProcessItem
                                WHERE RoutingProcessId = @RoutingProcessId
                                AND ItemNo = @ItemNo";
                        dynamicParameters.Add("RoutingProcessId", RoutingProcessId);
                        dynamicParameters.Add("ItemNo", ItemNo);

                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() > 0) throw new SystemException("【途程製程 + 項目】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.RoutingProcessItem (RoutingProcessId, ItemNo, ChkUnique
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@RoutingProcessId, @ItemNo, @ChkUnique
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoutingProcessId,
                                ItemNo,
                                ChkUnique,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

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

        #region //AddRipQcItem-- 新增途程製程加工參數 -- Ann 2022-09-06
        public string AddRipQcItem(string RipQcItemData)
        {
            try
            {
                if (RipQcItemData.Length <= 0) throw new SystemException("【加工參數】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        var ripQcItemJson = JObject.Parse(RipQcItemData);
                        int rowsAffected = 0;
                        foreach (var item in ripQcItemJson["qcItemInfo"])
                        {
                            #region //判斷加工流程卡資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.RoutingItemProcess
                                    WHERE ItemProcessId = @ItemProcessId";
                            dynamicParameters.Add("ItemProcessId", item["ItemProcessId"].ToString());

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("加工流程卡資料錯誤!");
                            #endregion

                            #region //判斷加工項目資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.QcItem
                                    WHERE QcItemId = @QcItemId";
                            dynamicParameters.Add("QcItemId", Convert.ToInt32(item["QcItemId"]));

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("加工項目資料錯誤!");
                            #endregion

                            #region //檢查是否已經有這項加工項目
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.RipQcItem
                                    WHERE ItemProcessId = @ItemProcessId
                                    AND QcItemId = @QcItemId";
                            dynamicParameters.Add("ItemProcessId", Convert.ToInt32(item["ItemProcessId"]));
                            dynamicParameters.Add("QcItemId", Convert.ToInt32(item["QcItemId"]));

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() > 0) break;
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.RipQcItem (ItemProcessId, QcItemId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@ItemProcessId, @QcItemId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ItemProcessId = item["ItemProcessId"].ToString(),
                                    QcItemId = Convert.ToInt32(item["QcItemId"]),
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
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

        #region //AddMoProcessChange-- 新增製令途程變更單單頭 -- Shintokuro 2022-10-27
        public string AddMoProcessChange(int MoId, int MpcUserId, string DocDate, string Remark)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //品異單單頭建立

                        #region //單號自動取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(MpcNo))), '000'), 3)) + 1 CurrentNum
                                FROM MES.MoProcessChange
								WHERE MpcNo NOT LIKE '%[A-Za-z]%'";

                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                        string DocDateNo = CreateDate.ToString("yyyyMM");
                        string MpcNo = DocDateNo + string.Format("{0:000}", currentNum);
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.MoProcessChange (CompanyId, MpcNo, MoId, MpcUserId, DocDate, Remark, Status
                                , CreateDate,LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.MpcId,INSERTED.MpcNo,INSERTED.MpcUserId
                                VALUES (@CompanyId, @MpcNo, @MoId, @MpcUserId, @DocDate, @Remark, @Status
                                , @CreateDate,@LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                MpcNo,
                                MoId,
                                MpcUserId,
                                DocDate,
                                Remark,
                                Status = "S",
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

        //#region //AddMpcRoutingProcess-- 新增變更單途程製程資料 -- Shintokuro 2022-11-04
        //public string AddMpcRoutingProcess(int MpcId, string ProcessData)
        //{
        //    try
        //    {
        //        if (ProcessData.Length <= 0) throw new SystemException("【製程別名】不能為空!");

        //        bool checkbool = false;
        //        using (TransactionScope transactionScope = new TransactionScope())
        //        {
        //            using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
        //            {
        //                #region //判斷製令途程變更單資料是否正確
        //                dynamicParameters = new DynamicParameters();
        //                sql = @"SELECT TOP 1 1
        //                        FROM MES.MoProcessChange
        //                        WHERE MpcId = @MpcId";
        //                dynamicParameters.Add("MpcId", MpcId);

        //                var result = sqlConnection.Query(sql, dynamicParameters);
        //                if (result.Count() <= 0) throw new SystemException("製令途程變更單資料錯誤!");
        //                #endregion

        //                var ProcessJson = JObject.Parse(ProcessData);
        //                int rowsAffected = 0;
        //                foreach (var item in ProcessJson["processInfo"])
        //                {
        //                    #region //判斷製程資料是否正確
        //                    dynamicParameters = new DynamicParameters();
        //                    sql = @"SELECT TOP 1 1
        //                            FROM MES.Process
        //                            WHERE ProcessId = @ProcessId";
        //                    dynamicParameters.Add("ProcessId", item["ProcessId"].ToString());

        //                    var result3 = sqlConnection.Query(sql, dynamicParameters);
        //                    if (result3.Count() <= 0) throw new SystemException("製程資料錯誤!");
        //                    #endregion

        //                    if (Convert.ToInt32(item["ProcessId"]) == 312) checkbool = true;

        //                    #region //取得最大排序數
        //                    int maxSortNumber = 0;
        //                    dynamicParameters = new DynamicParameters();
        //                    sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSortNumber
        //                            FROM MES.MpcRoutingProcess
        //                            WHERE MpcId = @MpcId";
        //                    dynamicParameters.Add("MpcId", MpcId);

        //                    var result4 = sqlConnection.Query(sql, dynamicParameters);

        //                    foreach (var item2 in result4)
        //                    {
        //                        maxSortNumber = Convert.ToInt32(item2.MaxSortNumber) + 1;
        //                    }
        //                    #endregion


        //                    dynamicParameters = new DynamicParameters();
        //                    sql = @"INSERT INTO MES.MpcRoutingProcess (MpcId, ProcessId, SortNumber, ProcessAlias, DisplayStatus, NecessityStatus, RandomStatus
        //                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
        //                            VALUES (@RoutingId, @ProcessId, @SortNumber, @ProcessAlias, @DisplayStatus, @NecessityStatus, @RandomStatus
        //                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
        //                    dynamicParameters.AddDynamicParams(
        //                        new
        //                        {
        //                            MpcId,
        //                            ProcessId = Convert.ToInt32(item["ProcessId"]),
        //                            SortNumber = maxSortNumber,
        //                            ProcessAlias = item["ProcessName"].ToString(),
        //                            DisplayStatus = "Y",
        //                            NecessityStatus = "Y",
        //                            RandomStatus = "Y",
        //                            CreateDate,
        //                            LastModifiedDate,
        //                            CreateBy,
        //                            LastModifiedBy
        //                        });

        //                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

        //                    rowsAffected += insertResult.Count();
        //                }

        //                #region //Response
        //                jsonResponse = JObject.FromObject(new
        //                {
        //                    status = "success",
        //                    msg = "(" + rowsAffected + " rows affected)"
        //                });
        //                #endregion
        //            }
        //            transactionScope.Complete();
        //        }

        //        if (checkbool == true)
        //        {
        //            using (TransactionScope transactionScope = new TransactionScope())
        //            {
        //                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
        //                {
        //                    #region //取得目前製程順序
        //                    dynamicParameters = new DynamicParameters();
        //                    sql = @"SELECT RoutingProcessId, SortNumber
        //                            FROM MES.RoutingProcess
        //                            WHERE RoutingId = @RoutingId
        //                            ORDER BY SortNumber";
        //                    dynamicParameters.Add("RoutingId", RoutingId);

        //                    var result = sqlConnection.Query(sql, dynamicParameters);

        //                    #region //先將全部製程順序改為-1
        //                    dynamicParameters = new DynamicParameters();
        //                    sql = @"UPDATE MES.RoutingProcess SET
        //                            SortNumber = SortNumber * -1,
        //                            LastModifiedDate = @LastModifiedDate,
        //                            LastModifiedBy = @LastModifiedBy
        //                            WHERE RoutingId = @RoutingId";
        //                    dynamicParameters.AddDynamicParams(
        //                        new
        //                        {
        //                            LastModifiedDate,
        //                            LastModifiedBy,
        //                            RoutingId
        //                        });

        //                    int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
        //                    #endregion

        //                    foreach (var item in result)
        //                    {
        //                        #region //先將製程順序全部+1
        //                        dynamicParameters = new DynamicParameters();
        //                        sql = @"UPDATE MES.RoutingProcess SET
        //                                SortNumber = @SortNumber,
        //                                LastModifiedDate = @LastModifiedDate,
        //                                LastModifiedBy = @LastModifiedBy
        //                                WHERE RoutingProcessId = @RoutingProcessId";
        //                        dynamicParameters.AddDynamicParams(
        //                            new
        //                            {
        //                                SortNumber = item.SortNumber + 1,
        //                                LastModifiedDate,
        //                                LastModifiedBy,
        //                                item.RoutingProcessId
        //                            });

        //                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
        //                        #endregion
        //                    }
        //                    #endregion

        //                    #region //新增退鎳站
        //                    dynamicParameters = new DynamicParameters();
        //                    sql = @"INSERT INTO MES.RoutingProcess (RoutingId, ProcessId, SortNumber, ProcessAlias, DisplayStatus, NecessityStatus, RandomStatus
        //                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
        //                            VALUES (@RoutingId, @ProcessId, @SortNumber, @ProcessAlias, @DisplayStatus, @NecessityStatus, @RandomStatus
        //                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
        //                    dynamicParameters.AddDynamicParams(
        //                        new
        //                        {
        //                            RoutingId,
        //                            ProcessId = 313,
        //                            SortNumber = 1,
        //                            ProcessAlias = "退鎳",
        //                            DisplayStatus = "N",
        //                            NecessityStatus = "N",
        //                            RandomStatus = "N",
        //                            CreateDate,
        //                            LastModifiedDate,
        //                            CreateBy,
        //                            LastModifiedBy
        //                        });

        //                    var insertResult = sqlConnection.Query(sql, dynamicParameters);
        //                    #endregion
        //                }
        //                transactionScope.Complete();
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        #region //Response
        //        jsonResponse = JObject.FromObject(new
        //        {
        //            status = "errorForDA",
        //            msg = e.Message
        //        });
        //        #endregion

        //        logger.Error(e.Message);
        //    }

        //    return jsonResponse.ToString();
        //}
        //#endregion

        #region //AddMpcRoutingProcessAll-- 新增製令途程變更單綁定製程 -- Shintokuro 2022-10-31
        public string AddMpcRoutingProcessAll(int MpcId, int MoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷製令途程變更單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.MoProcessChange
                                WHERE MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製令途程變更單資料錯誤!");
                        #endregion

                        #region //撈取製令下的製程
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MoProcessId,a.MoId,a.SortNumber,a.ProcessAlias,a.RoutingItemProcessDesc,a.Remark,a.DisplayStatus,a.NecessityStatus
                                ,a.RandomStatus,a.OspFlag,a.ProcessCheckStatus,a.ProcessCheckType
                                FROM MES.MoProcess a
                                WHERE a.MoId = @MoId
                                ";
                        dynamicParameters.Add("MoId", MoId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製令ID有問題找不到資料");
                        #endregion

                        int rowsAffected = 0;
                        foreach (var item in result)
                        {
                            #region //判斷製程資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.MoProcess
                                    WHERE MoProcessId = @MoProcessId";
                            dynamicParameters.Add("MoProcessId", item.MoProcessId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("製程資料錯誤!");
                            #endregion

                            #region //判斷製程是否有重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 ProcessAlias
                                    FROM MES.MpcRoutingProcess
                                    WHERE MoProcessId = @MoProcessId
                                    AND MpcId = @MpcId";
                            dynamicParameters.Add("MoProcessId", item.MoProcessId);
                            dynamicParameters.Add("MpcId", MpcId);

                            result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() > 0)
                            {
                                string ProcessAlias = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ProcessAlias;
                                throw new SystemException("【製程:" + ProcessAlias + "】重複，請重新選擇!");
                            }
                            #endregion

                            #region //新增SQL
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.MpcRoutingProcess (MpcId, MoProcessId, SortNumber 
                                    , ProcessAlias, RoutingItemProcessDesc, Remark, DisplayStatus
                                    , NecessityStatus, RandomStatus, OspFlag, ProcessCheckStatus, ProcessCheckType
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@MpcId, @MoProcessId, @SortNumber 
                                    , @ProcessAlias, @RoutingItemProcessDesc, @Remark, @DisplayStatus
                                    , @NecessityStatus, @RandomStatus, @OspFlag, @ProcessCheckStatus, @ProcessCheckType
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MpcId,
                                    MoProcessId = item.MoProcessId,
                                    SortNumber = item.SortNumber,
                                    ProcessAlias = item.ProcessAlias,
                                    RoutingItemProcessDesc = item.RoutingItemProcessDesc,
                                    Remark = item.Remark,
                                    DisplayStatus = item.DisplayStatus,
                                    NecessityStatus = item.NecessityStatus,
                                    RandomStatus = item.RandomStatus,
                                    OspFlag = item.OspFlag,
                                    ProcessCheckStatus = item.ProcessCheckStatus,
                                    ProcessCheckType = item.ProcessCheckType,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion

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

        #region //AddMpcBarcode-- 新增製令途程變更單綁定條碼 -- Shintokuro 2022-10-31
        public string AddMpcBarcode(int MpcId, string BarcodeIds)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (MpcId <= 0) throw new SystemException("【變更單ID】不能為空!");
                        if (BarcodeIds.Length <= 0) throw new SystemException("【條碼清單】不能為空!");
                        int rowsAffected = 0;

                        dynamicParameters = new DynamicParameters();
                        #region //判斷製令途程變更單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MoId
                                FROM MES.MoProcessChange
                                WHERE MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製令途程變更單資料錯誤!");
                        int MoId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MoId;
                        #endregion


                        foreach (var BarcodeNo in BarcodeIds.Split(','))
                        {
                            #region //判斷條碼是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.ManufactureOrder a
                                    LEFT JOIN MES.MrDetail b on a.MoId = b.MoId
                                    LEFT JOIN MES.MrBarcodeRegister b1 on b.MrDetailId = b1.MrDetailId
                                    LEFT JOIN MES.MoSetting c on a.MoId =c.MoId
                                    LEFT JOIN MES.BarcodePrint c1 on c.MoSettingId =c1.MoSettingId
                                    WHERE b1.BarcodeNo = @BarcodeNo or c1.BarcodeNo = @BarcodeNo 
                                    AND a.MoId = @MoId";
                            dynamicParameters.Add("BarcodeNo", BarcodeNo);
                            dynamicParameters.Add("MoId", MoId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("條碼資料錯誤!");
                            #endregion

                            #region //判斷條碼是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.MpcBarcode a
                                    WHERE a.BarcodeNo = @BarcodeNo
                                    AND a.MpcId = @MpcId";
                            dynamicParameters.Add("BarcodeNo", BarcodeNo);
                            dynamicParameters.Add("MpcId", MpcId);

                            result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() > 0) throw new SystemException("【" + BarcodeNo + "】條碼重複了,不可以新增!");
                            #endregion

                            #region //新增SQL
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.MpcBarcode (MpcId, BarcodeNo
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@MpcId, @BarcodeNo
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MpcId,
                                    BarcodeNo,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion

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

        #region// AddAiRoutingInfo -- 自動流程卡生成 --MarkChen 2023.08.21




        #endregion

        #endregion

        #region //Update
        #region //UpdateRouting -- 更新途程資料 -- Ann 2022-07-14
        public string UpdateRouting(int RoutingId, string RoutingName, string RoutingType, int ModeId)
        {
            try
            {
                if (RoutingName.Length <= 0) throw new SystemException("【途程名稱】不能為空!");
                if (RoutingName.Length > 100) throw new SystemException("【途程名稱】長度錯誤!");
                if (!Regex.IsMatch(RoutingType, "^(1|2)$", RegexOptions.IgnoreCase)) throw new SystemException("【途程類型】錯誤!");
                if (ModeId <= 0) throw new SystemException("【生產模式】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷途程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT RoutingName, RoutingType, ModeId, RoutingConfirm
                                FROM MES.Routing
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程資料錯誤!");

                        foreach (var item in result)
                        {
                            if (item.RoutingType != RoutingType || item.ModeId != ModeId)
                            {
                                if (item.RoutingConfirm != "N") throw new SystemException("途程已被確認，無法更改!");
                            }
                        }
                        #endregion

                        #region //判斷生產模式是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdMode
                                WHERE ModeId = @ModeId";
                        dynamicParameters.Add("ModeId", ModeId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("生產模式資料錯誤!");
                        #endregion

                        #region //判斷 生產模式+途程名稱 是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE ModeId = @ModeId
                                AND RoutingName = @RoutingName
                                AND RoutingId != @RoutingId";
                        dynamicParameters.Add("ModeId", ModeId);
                        dynamicParameters.Add("RoutingName", RoutingName);
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【生產模式 + 途程名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Routing SET
                                RoutingName = @RoutingName,
                                RoutingType = @RoutingType,
                                ModeId = @ModeId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoutingName,
                                RoutingType,
                                ModeId,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoutingId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateRoutingStatus -- 更新途程啟用狀態 -- Ann 2022-07-14
        public string UpdateRoutingStatus(int RoutingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷途程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status, RoutingConfirm
                                FROM MES.Routing
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程資料錯誤!");

                        foreach (var item in result)
                        {
                            //if (item.RoutingConfirm != "N") throw new SystemException("途程已被確認，無法更改!");
                        }
                        #endregion

                        #region //調整為相反狀態
                        string Status = "";
                        foreach (var item in result)
                        {
                            Status = item.Status;
                        }

                        switch (Status)
                        {
                            case "A":
                                Status = "S";
                                break;
                            case "S":
                                Status = "A";
                                break;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Routing SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoutingId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateRoutingProcess -- 更新途程製程資料 -- Ann 2022-07-15
        public string UpdateRoutingProcess(int RoutingProcessId, int RoutingId, int ProcessId, string ProcessAlias
            , string DisplayStatus, string NecessityStatus, string ProcessCheckStatus, string ProcessCheckType, string PackageFlag, string ConsumeFlag)
        {
            try
            {
                if (ProcessId <= 0) throw new SystemException("【製程資訊】不能為空!");
                if (ProcessAlias.Length <= 0) throw new SystemException("【製程別名】不能為空!");
                if (ProcessAlias.Length > 32) throw new SystemException("【製程別名】長度錯誤!");
                if (!Regex.IsMatch(NecessityStatus, "^(Y|N)$", RegexOptions.IgnoreCase)) throw new SystemException("【必要過站】錯誤!");
                if (!Regex.IsMatch(DisplayStatus, "^(Y|N)$", RegexOptions.IgnoreCase)) throw new SystemException("【是否顯示在流程卡】錯誤!");
                if (ProcessCheckStatus.Length <= 0) throw new SystemException("【是否支援工程檢】不能為空!");
                if (ProcessCheckStatus == "Y" && ProcessCheckType.Length <= 0) throw new SystemException("【檢驗頻率】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷途程製程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.RoutingProcess
                                WHERE RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程製程資料錯誤!");
                        #endregion

                        #region //判斷途程是否已被確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE RoutingConfirm = @RoutingConfirm
                                AND RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingConfirm", "Y");
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("此途程已被確認，無法更動製程資料!");
                        #endregion

                        #region //判斷途程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("途程資料錯誤!");
                        #endregion

                        #region //判斷製程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Process
                                WHERE ProcessId = @ProcessId";
                        dynamicParameters.Add("ProcessId", ProcessId);

                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() <= 0) throw new SystemException("製程資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.RoutingProcess SET
                                RoutingId = @RoutingId,
                                ProcessId = @ProcessId,
                                ProcessAlias = @ProcessAlias,
                                DisplayStatus = @DisplayStatus,
                                NecessityStatus = @NecessityStatus,
                                ProcessCheckStatus = @ProcessCheckStatus,
                                ProcessCheckType = @ProcessCheckType,
                                PackageFlag = @PackageFlag,
ConsumeFlag  = @ConsumeFlag,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoutingId,
                                ProcessId,
                                ProcessAlias,
                                DisplayStatus,
                                NecessityStatus,
                                ProcessCheckStatus,
                                ProcessCheckType = ProcessCheckType.Length > 0 ? ProcessCheckType : null,
                                PackageFlag,
                                ConsumeFlag,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoutingProcessId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateNecessityStatus -- 更新必要過站狀態 -- Ann 2022-07-15
        public string UpdateNecessityStatus(int RoutingProcessId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷途程製程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 NecessityStatus, RoutingId
                                FROM MES.RoutingProcess
                                WHERE RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程製程資料錯誤!");

                        int RoutingId = -1;
                        string NecessityStatus = "";
                        foreach (var item in result)
                        {
                            RoutingId = item.RoutingId;
                            NecessityStatus = item.NecessityStatus;
                        }
                        #endregion

                        #region //判斷途程是否已被確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE RoutingConfirm = @RoutingConfirm
                                AND RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingConfirm", "Y");
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("此途程已被確認，無法更動製程資料!");
                        #endregion

                        #region //調整為相反狀態
                        switch (NecessityStatus)
                        {
                            case "Y":
                                NecessityStatus = "N";
                                break;
                            case "N":
                                NecessityStatus = "Y";
                                break;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.RoutingProcess SET
                                NecessityStatus = @NecessityStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                NecessityStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoutingProcessId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateRoutingProcessSort -- 更新途程製程排序 -- Ann 2022-07-18
        public string UpdateRoutingProcessSort(string RoutingProcessSort)
        {
            try
            {
                if (RoutingProcessSort.Length <= 0) throw new SystemException("【製程順序】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        string[] RoutingProcessIdArray = RoutingProcessSort.Split(',');
                        int RoutingId = 0;
                        #region //判斷途程製程資料是否正確
                        foreach (var item in RoutingProcessIdArray)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 RoutingId
                                FROM MES.RoutingProcess
                                WHERE RoutingProcessId = @RoutingProcessId";
                            dynamicParameters.Add("RoutingProcessId", item);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("途程製程資料錯誤!");

                            foreach (var item2 in result)
                            {
                                RoutingId = Convert.ToInt32(item2.RoutingId);
                            }

                            #region //判斷途程是否已被確認
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE RoutingConfirm = @RoutingConfirm
                                AND RoutingId = @RoutingId";
                            dynamicParameters.Add("RoutingConfirm", "Y");
                            dynamicParameters.Add("RoutingId", RoutingId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() > 0) throw new SystemException("此途程已被確認，無法更動製程資料!");
                            #endregion
                        }
                        #endregion

                        #region //先將此途程製程排序改為-1
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.RoutingProcess SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新順序
                        int sort = 1;
                        foreach (var item3 in RoutingProcessIdArray)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.RoutingProcess SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE RoutingProcessId = @RoutingProcessId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = sort,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    RoutingProcessId = Convert.ToInt32(item3)
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            sort++;
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

        #region //UpdateRoutingItem -- 更新途程品號資料 -- Ann 2022-07-19
        public string UpdateRoutingItem(int RoutingItemId, int RoutingId, int ControlId, int MtlItemId, string Status)
        {
            try
            {
                if (!Regex.IsMatch(Status, "^(A|S)$", RegexOptions.IgnoreCase)) throw new SystemException("【必要過站】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷途程品號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MtlItemId OriMtlItemId
                                FROM MES.RoutingItem
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程品號資料錯誤!");

                        int OriMtlItemId = -1;
                        foreach (var item in result)
                        {
                            OriMtlItemId = item.OriMtlItemId;
                        }
                        #endregion

                        #region //判斷途程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("途程資料錯誤!");
                        #endregion

                        #region //判斷研發設計圖版控資料是否正確
                        if (ControlId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM PDM.RdDesignControl
                                WHERE ControlId = @ControlId";
                            dynamicParameters.Add("ControlId", ControlId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("研發設計圖版控資料錯誤!");
                        }
                        #endregion

                        #region //判斷品號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlItem
                                WHERE MtlItemId = @MtlItemId";
                        dynamicParameters.Add("MtlItemId", MtlItemId);

                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() <= 0) throw new SystemException("品號資料錯誤!");
                        #endregion

                        #region //判斷此途程品號是否已被製令綁定
                        if (MtlItemId != OriMtlItemId)
                        {
                            sql = @"SELECT TOP 1 1
                                    FROM MES.MoRouting a
                                    WHERE a.RoutingItemId = @RoutingItemId";
                            dynamicParameters.Add("RoutingItemId", RoutingItemId);

                            var result5 = sqlConnection.Query(sql, dynamicParameters);
                            if (result5.Count() > 0) throw new SystemException("此途程品號已被製令綁定，無法刪除!");
                        }
                        #endregion

                        //#region //判斷 途程+研發設計圖版控 是否重複
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT TOP 1 1
                        //        FROM MES.RoutingItem
                        //        WHERE RoutingId = @RoutingId
                        //        AND ControlId = @ControlId
                        //        AND RoutingItemId != @RoutingItemId";
                        //dynamicParameters.Add("RoutingId", RoutingId);
                        //dynamicParameters.Add("ControlId", ControlId);
                        //dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        //var result5 = sqlConnection.Query(sql, dynamicParameters);
                        //if (result5.Count() > 0) throw new SystemException("【途程 + 研發設計圖版控】重複，請重新輸入!");
                        //#endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.RoutingItem SET
                                RoutingId = @RoutingId,
                                ControlId = @ControlId,
                                MtlItemId = @MtlItemId,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoutingId,
                                ControlId = ControlId > 0 ? (int?)ControlId : null,
                                MtlItemId,
                                Status,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoutingItemId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //重綁加工屬性
                        #region //先刪除原本的屬性
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.RoutingItemAttribute
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //塞加工屬性
                        #region //找此張設計圖加工屬性
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.AttributeId, a.AttributeItemId, a.AttributeValue
                                , a.UpperLimit, a.LowerLimit
                                FROM PDM.RdDesignAttribute a
                                WHERE a.ControlId = @ControlId";
                        dynamicParameters.Add("ControlId", ControlId);

                        var result6 = sqlConnection.Query(sql, dynamicParameters);
                        if (result6.Count() > 0)
                        {
                            foreach (var item in result6)
                            {
                                //開始塞資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.RoutingItemAttribute (RoutingItemId, AttributeItemId
                                        , AttributeValue, UpperLimit, LowerLimit
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@RoutingItemId, @AttributeItemId
                                        , @AttributeValue, @UpperLimit, @LowerLimit
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RoutingItemId,
                                        item.AttributeItemId,
                                        item.AttributeValue,
                                        item.UpperLimit,
                                        item.LowerLimit,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
                        }
                        #endregion
                        #endregion

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

        #region //UpdateRoutingItemStatus -- 更新途程品號狀態 -- Ann 2022-07-20
        public string UpdateRoutingItemStatus(int RoutingItemId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷途程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.RoutingItem
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程品號資料錯誤!");
                        #endregion

                        #region //判斷此途程品號是否已被製令綁定
                        sql = @"SELECT TOP 1 1
                                FROM MES.MoRouting a
                                WHERE a.RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        //if (result2.Count() > 0) throw new SystemException("此途程品號已被製令綁定，無法刪除!");
                        #endregion

                        #region //調整為相反狀態
                        string Status = "";
                        foreach (var item in result)
                        {
                            Status = item.Status;
                        }

                        switch (Status)
                        {
                            case "A":
                                Status = "S";
                                break;
                            case "S":
                                Status = "A";
                                break;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.RoutingItem SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoutingItemId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateRoutingItemAttribute-- 更新途程品號加工屬性 -- Ann 2022-07-21
        public string UpdateRoutingItemAttribute(string AttributeData)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (AttributeData.Length <= 0) throw new SystemException("【屬性資料】不能為空!");

                        var AttributeJson = JObject.Parse(AttributeData);
                        dynamicParameters = new DynamicParameters();
                        int rowsAffected = 0;
                        foreach (var item in AttributeJson["attributeInfo"])
                        {
                            #region //判斷途程品號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.RoutingItem a
                                    WHERE a.RoutingItemId = @RoutingItemId";
                            dynamicParameters.Add("RoutingItemId", item["RoutingItemId"].ToString());

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("途程品號資料錯誤!");
                            #endregion

                            #region //判斷加工屬性資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.RdAttributeItem a
                                    WHERE a.AttributeItemId = @AttributeItemId";
                            dynamicParameters.Add("AttributeItemId", item["AttributeItemId"].ToString());

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("加工屬性資料錯誤!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.RoutingItemAttribute SET
                                    RoutingItemId = @RoutingItemId,
                                    AttributeItemId = @AttributeItemId,
                                    AttributeValue = @AttributeValue,
                                    UpperLimit = @UpperLimit,
                                    LowerLimit = @LowerLimit,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE AttributeId = @AttributeId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RoutingItemId = item["RoutingItemId"].ToString(),
                                    AttributeItemId = item["AttributeItemId"].ToString(),
                                    AttributeValue = item["AttributeValue"].ToString(),
                                    UpperLimit = item["UpperLimit"].ToString(),
                                    LowerLimit = item["LowerLimit"].ToString(),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    AttributeId = item["AttributeId"].ToString()
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //ConfirmRouting -- 確認途程 -- Ann 2022-07-25
        public string UpdateRoutingConfirm(int RoutingId)
        {
            try
            {
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷途程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程資料錯誤!");
                        #endregion

                        #region //判斷此途程是否至少有一個製程
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.RoutingProcess
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("此途程尚未建立任何製程!");
                        #endregion

                        #region //同步新增加工細項製程
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RoutingItemId
                                , b.MtlItemName
                                FROM MES.RoutingItem a 
                                INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                WHERE a.RoutingId = @RoutingId
                                AND a.RoutingItemConfirm = 'Y'";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var RoutingItemResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in RoutingItemResult)
                        {
                            #region //取得新增的製程項目
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.RoutingProcessId, a.RoutingId, a.ProcessId, a.ProcessAlias
                                    , ISNULL(b.ProcessTemplateDesc, '') ProcessTemplateDesc
                                    , ISNULL(b.Remark, '') Remark
                                    , c.ProcessNo, c.ProcessName
                                    , d.ModeId
                                    FROM MES.RoutingProcess a 
                                    LEFT JOIN MES.RoutingProcessTemplate b ON a.ProcessId = b.ProcessId
                                    INNER JOIN MES.Process c ON a.ProcessId = c.ProcessId
                                    INNER JOIN MES.Routing d ON a.RoutingId = d.RoutingId
                                    WHERE a.RoutingId = @RoutingId
                                    AND a.RoutingProcessId NOT IN (
                                        SELECT x.RoutingProcessId
                                        FROM MES.RoutingItemProcess x 
                                        INNER JOIN MES.RoutingItem xa ON x.RoutingItemId = xa.RoutingItemId
                                        WHERE xa.RoutingItemId = @RoutingItemId
                                        AND xa.RoutingItemConfirm = 'Y'
                                    )";
                            dynamicParameters.Add("RoutingId", RoutingId);
                            dynamicParameters.Add("RoutingItemId", item.RoutingItemId);

                            var RoutingProcessResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item2 in RoutingProcessResult)
                            {
                                string ProcessTemplateDesc = "";
                                if (item2.ProcessId == 11) ProcessTemplateDesc = item.MtlItemName;

                                #region //Insert MES.RoutingItemProcess
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.RoutingItemProcess (RoutingItemId, RoutingProcessId, RoutingItemProcessDesc, Remark, DisplayStatus
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.ItemProcessId
                                        VALUES (@RoutingItemId, @RoutingProcessId, @RoutingItemProcessDesc, @Remark, @DisplayStatus
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        item.RoutingItemId,
                                        item2.RoutingProcessId,
                                        RoutingItemProcessDesc = ProcessTemplateDesc,
                                        item2.Remark,
                                        DisplayStatus = "Y",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                int ItemProcessId = -1;
                                foreach (var item3 in insertResult)
                                {
                                    ItemProcessId = item3.ItemProcessId;
                                }
                                #endregion

                                #region //確認製程參數是否有量測項目
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.QcItemId , a.DesignValue, a.UpperTolerance, a.LowerTolerance, a.Remark, a.BallMark, a.Unit, a.QmmDetailId, a.QcItemDesc
                                        FROM MES.ProcessParameterQcItem a 
                                        INNER JOIN MES.ProcessParameter b ON a.ParameterId = b.ParameterId
                                        WHERE b.ProcessId = @ProcessId 
                                        AND b.ModeId = @ModeId";
                                dynamicParameters.Add("ProcessId", item2.ProcessId);
                                dynamicParameters.Add("ModeId", item2.ModeId);

                                var ProcessParameterQcItemResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item3 in ProcessParameterQcItemResult)
                                {
                                    #region //Insert RoutingItemQcItem
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.RoutingItemQcItem (ItemProcessId, RoutingItemId, QcItemId, QcItemDesc, DesignValue, UpperTolerance, LowerTolerance, Remark
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy, BallMark, Unit, QmmDetailId)
                                            OUTPUT INSERTED.RoutingItemId
                                            VALUES (@ItemProcessId, @RoutingItemId, @QcItemId, @QcItemDesc, @DesignValue, @UpperTolerance, @LowerTolerance, @Remark
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @BallMark, @Unit, @QmmDetailId)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            ItemProcessId,
                                            item.RoutingItemId,
                                            item3.QcItemId,
                                            item3.QcItemDesc,
                                            item3.DesignValue,
                                            item3.UpperTolerance,
                                            item3.LowerTolerance,
                                            item3.Remark,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy,
                                            item3.BallMark,
                                            item3.Unit,
                                            item3.QmmDetailId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                            }
                            #endregion
                        }
                        #endregion

                        #region //Update Routing RoutingConfirm
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Routing SET
                                RoutingConfirm = @RoutingConfirm,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoutingConfirm = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                RoutingId
                            });

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

        #region //UnbindConfirmRouting -- 反確認途程 -- Ann 2022-08-05
        public string UnbindConfirmRouting(int RoutingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷途程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程資料錯誤!");
                        #endregion

                        #region //判斷此途程是否已建立途程品號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.RoutingItem
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        //if (result2.Count() > 0) throw new SystemException("此途程已建立途程品號，無法解除確認狀態!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Routing SET
                                RoutingConfirm = @RoutingConfirm,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoutingConfirm = "N",
                                LastModifiedDate,
                                LastModifiedBy,
                                RoutingId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateRoutingItemConfirm -- 確認途程品號 -- Ann 2022-07-25
        public string UpdateRoutingItemConfirm(int RoutingItemId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷途程品號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RoutingId
                                FROM MES.RoutingItem a
                                WHERE a.RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程品號資料錯誤!");

                        int RoutingId = 0;
                        foreach (var item in result)
                        {
                            RoutingId = item.RoutingId;
                        }
                        #endregion

                        #region //判斷此途程是否已經確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE RoutingId = @RoutingId
                                AND RoutingConfirm = @RoutingConfirm";
                        dynamicParameters.Add("RoutingId", RoutingId);
                        dynamicParameters.Add("RoutingConfirm", "Y");

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("此途程尚未確認!");
                        #endregion

                        #region //Update RoutingItem
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.RoutingItem SET
                                RoutingItemConfirm = @RoutingItemConfirm,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoutingItemConfirm = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                RoutingItemId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //將加工細節填入MES.RoutingItemProcess
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RoutingProcessId, a.RoutingId, a.ProcessId, a.ProcessAlias
                                , ISNULL(b.ProcessTemplateDesc, '') ProcessTemplateDesc
                                , ISNULL(b.Remark, '') Remark
                                , c.ProcessNo, c.ProcessName
                                , e.MtlItemName
                                , f.ModeId
                                FROM MES.RoutingProcess a
                                LEFT JOIN MES.RoutingProcessTemplate b ON a.ProcessId = b.ProcessId
                                INNER JOIN MES.Process c ON a.ProcessId = c.ProcessId
                                INNER JOIN MES.RoutingItem d ON a.RoutingId = d.RoutingId
                                INNER JOIN PDM.MtlItem e ON d.MtlItemId = e.MtlItemId
                                INNER JOIN MES.Routing f ON a.RoutingId = f.RoutingId
                                WHERE a.RoutingId = @RoutingId
                                AND d.RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingId", RoutingId);
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);

                        int RoutingProcessId = 0;
                        string ProcessTemplateDesc = "";
                        string Remark = "";
                        foreach (var item in result3)
                        {
                            int ItemProcessId = -1;

                            RoutingProcessId = item.RoutingProcessId;
                            ProcessTemplateDesc = item.ProcessTemplateDesc;
                            Remark = item.Remark;

                            if (item.ProcessId == 11) ProcessTemplateDesc = item.MtlItemName;

                            #region //Insert MES.RoutingItemProcess

                            #region //確認此途程品號已建立的加工細節
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ItemProcessId
                                    FROM MES.RoutingItemProcess a
                                    WHERE a.RoutingItemId = @RoutingItemId
                                    AND a.RoutingProcessId = @RoutingProcessId";
                            dynamicParameters.Add("RoutingItemId", RoutingItemId);
                            dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                            var CheckRoutingItemProcessResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item2 in CheckRoutingItemProcessResult)
                            {
                                ItemProcessId = item2.ItemProcessId;
                            }
                            #endregion

                            if (CheckRoutingItemProcessResult.Count() <= 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.RoutingItemProcess (RoutingItemId, RoutingProcessId, RoutingItemProcessDesc, Remark, DisplayStatus
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.ItemProcessId
                                        VALUES (@RoutingItemId, @RoutingProcessId, @RoutingItemProcessDesc, @Remark, @DisplayStatus
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RoutingItemId,
                                        RoutingProcessId,
                                        RoutingItemProcessDesc = ProcessTemplateDesc,
                                        Remark,
                                        DisplayStatus = "Y",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                foreach (var item2 in insertResult)
                                {
                                    ItemProcessId = item2.ItemProcessId;
                                }
                            }
                            #endregion

                            #region //將製程參數設定之量測參數帶進來
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcItemId, a.DesignValue, a.UpperTolerance, a.LowerTolerance, a.Remark, a.QcItemDesc, a.BallMark, a.Unit, a.QmmDetailId
                                    FROM MES.ProcessParameterQcItem a 
                                    INNER JOIN MES.ProcessParameter b ON a.ParameterId = b.ParameterId
                                    WHERE b.ProcessId = @ProcessId 
                                    AND b.ModeId = @ModeId";
                            dynamicParameters.Add("ProcessId", item.ProcessId);
                            dynamicParameters.Add("ModeId", item.ModeId);

                            var ProcessParameterQcItemResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item2 in ProcessParameterQcItemResult)
                            {
                                #region //Insert MES.RoutingItemQcItem
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.RoutingItemQcItem (RoutingItemId, ItemProcessId, QcItemId, DesignValue, UpperTolerance, LowerTolerance, Remark, QcItemDesc
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy, BallMark, Unit, QmmDetailId)
                                        OUTPUT INSERTED.ItemProcessId
                                        VALUES (@RoutingItemId, @ItemProcessId, @QcItemId, @DesignValue, @UpperTolerance, @LowerTolerance, @Remark, @QcItemDesc
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @BallMark, @Unit, @QmmDetailId)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RoutingItemId,
                                        ItemProcessId,
                                        item2.QcItemId,
                                        item2.DesignValue,
                                        item2.UpperTolerance,
                                        item2.LowerTolerance,
                                        item2.Remark,
                                        item2.QcItemDesc,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy,
                                        item2.BallMark,
                                        item2.Unit,
                                        item2.QmmDetailId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion
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

        #region //UpdateRoutingItemProcess -- 更新加工流程卡內容 -- Ann 2022-07-25
        public string UpdateRoutingItemProcess(string RoutingItemProcessData)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        JObject routingItemProcessJson = JObject.Parse(RoutingItemProcessData);
                        int rowsAffected = 0;
                        foreach (var item in routingItemProcessJson["routingItemProcessInfo"])
                        {
                            if (item["RoutingItemProcessDesc"].ToString().IndexOf("$") != -1) throw new SystemException("尚有加工係數尚未維護!");
                            //if (Convert.ToInt32(item["CycleTime"]) <= 0) throw new SystemException("預計完工時間不能為0或小於0!");
                            //if (Convert.ToInt32(item["MoveTime"]) <= 0) throw new SystemException("預計移動到下製程時間不能為0或小於0!");

                            #region //判斷途程品號是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.RoutingItem
                                    WHERE RoutingItemId = @RoutingItemId";
                            dynamicParameters.Add("RoutingItemId", item["RoutingItemId"].ToString());

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("途程品號資料錯誤!");
                            #endregion

                            #region //判斷流程卡資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.RoutingItemProcess
                                    WHERE ItemProcessId = @ItemProcessId";
                            dynamicParameters.Add("ItemProcessId", item["ItemProcessId"].ToString());

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("流程卡加工資料錯誤!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.RoutingItemProcess SET
                                    RoutingItemProcessDesc = @RoutingItemProcessDesc,
                                    Remark = @Remark,
                                    AttrSetting = @AttrSetting,
                                    CycleTime = @CycleTime,
                                    MoveTime = @MoveTime,
                                    ProcessingTime = @ProcessingTime,
                                    WatingTime = @WatingTime,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ItemProcessId = @ItemProcessId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RoutingItemProcessDesc = item["RoutingItemProcessDesc"].ToString(),
                                    Remark = item["Remark"].ToString(),
                                    AttrSetting = item["AttrSetting"].ToString(),
                                    CycleTime = Convert.ToInt32(item["CycleTime"]),
                                    MoveTime = Convert.ToInt32(item["MoveTime"]),
                                    ProcessingTime = Convert.ToInt32(item["ProcessingTime"]),
                                    WatingTime = Convert.ToInt32(item["WatingTime"]),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ItemProcessId = item["ItemProcessId"].ToString()
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateDisplayStatus -- 更新是否顯示在流程卡狀態 -- Ann 2022-08-12
        public string UpdateDisplayStatus(int RoutingProcessId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷途程製程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DisplayStatus, RoutingId
                                FROM MES.RoutingProcess
                                WHERE RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程製程資料錯誤!");

                        int RoutingId = -1;
                        string DisplayStatus = "";
                        foreach (var item in result)
                        {
                            RoutingId = item.RoutingId;
                            DisplayStatus = item.DisplayStatus;
                        }
                        #endregion

                        #region //判斷途程是否已被確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE RoutingConfirm = @RoutingConfirm
                                AND RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingConfirm", "Y");
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("此途程已被確認，無法更動製程資料!");
                        #endregion

                        #region //調整為相反狀態
                        switch (DisplayStatus)
                        {
                            case "Y":
                                DisplayStatus = "N";
                                break;
                            case "N":
                                DisplayStatus = "Y";
                                break;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.RoutingProcess SET
                                DisplayStatus = @DisplayStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DisplayStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoutingProcessId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateRoutingProcessItem -- 更新途程製程屬性檔 -- Ann 2022-08-12
        public string UpdateRoutingProcessItem(int RoutingProcessItemId, int RoutingProcessId, string ItemNo, string ChkUnique)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ChkUnique.Length <= 0) throw new SystemException("【是否檢核重複值】不能為空!");

                        dynamicParameters = new DynamicParameters();
                        #region //判斷途程製程屬性檔資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.RoutingProcessItem a
                                WHERE a.RoutingProcessItemId = @RoutingProcessItemId";
                        dynamicParameters.Add("RoutingProcessItemId", RoutingProcessItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程製程屬性檔資料錯誤!");
                        #endregion

                        #region //判斷途程製程資料是否正確
                        sql = @"SELECT TOP 1 RoutingId
                                FROM MES.RoutingProcess a
                                WHERE a.RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("途程製程資料錯誤!");

                        int RoutingId = 0;
                        foreach (var item in result2)
                        {
                            RoutingId = item.RoutingId;
                        }
                        #endregion

                        #region //判斷類別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Type a
                                WHERE a.TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeNo", ItemNo);

                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() <= 0) throw new SystemException("類別資料錯誤!");
                        #endregion

                        #region //判斷 途程製程+項目 是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.RoutingProcessItem
                                WHERE RoutingProcessId = @RoutingProcessId
                                AND ItemNo = @ItemNo
                                AND RoutingProcessItemId != @RoutingProcessItemId";
                        dynamicParameters.Add("RoutingProcessId", RoutingProcessId);
                        dynamicParameters.Add("ItemNo", ItemNo);
                        dynamicParameters.Add("RoutingProcessItemId", RoutingProcessItemId);

                        var result5 = sqlConnection.Query(sql, dynamicParameters);
                        if (result5.Count() > 0) throw new SystemException("【途程製程 + 項目】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.RoutingProcessItem SET
                                RoutingProcessId = @RoutingProcessId,
                                ItemNo = @ItemNo,
                                ChkUnique = @ChkUnique,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoutingProcessItemId = @RoutingProcessItemId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoutingProcessId,
                                ItemNo,
                                ChkUnique,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoutingProcessItemId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateMpcIdStatus -- 更新製令途程變更單狀態 -- Shintokuro 2022.10.27
        public string UpdateMpcIdStatus(int MpcId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();


                        #region //判斷生產單元是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM MES.MoProcessChange
                                WHERE MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

                        var resultStatus = sqlConnection.Query(sql, dynamicParameters);
                        if (resultStatus.Count() <= 0) throw new SystemException("製令途程變更單資料錯誤!");
                        var Status = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).Status;
                        if (Status == "A") throw new SystemException("變更單已啟用,不可以更動!!!");

                        #region //判斷啟用時變更單內的條碼條碼是否有開工
                        sql = @"SELECT c.BarcodeNo,b.ProcessAlias
                                FROM MES.MoProcessChange a
                                INNER JOIN MES.MpcRoutingProcess b ON a.MpcId = b.MpcId
                                INNER JOIN MES.MpcBarcode c ON a.MpcId = c.MpcId
                                AND EXISTS(SELECT 1
                                           FROM MES.MoProcess d
			                               WHERE b.MoProcessId = d.MoProcessId
			                               AND b.SortNumber != d.SortNumber
				                           AND EXISTS(SELECT 1
				                                      FROM MES.BarcodeProcess e
							                          INNER JOIN MES.Barcode f on e.BarcodeId = f.BarcodeId
							                          WHERE d.MoProcessId = e.MoProcessId
							                          AND f.BarcodeNo = c.BarcodeNo
                                                      )
                                           )
                                WHERE a.MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            var ProcessAlias = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ProcessAlias;
                            string BarcodeNo = "";
                            int itemNum = 0;
                            foreach (var item in result)
                            {
                                if (itemNum == 0)
                                {
                                    BarcodeNo += "\n" + item.BarcodeNo;
                                }
                                else
                                {
                                    BarcodeNo += "\n," + item.BarcodeNo;

                                }
                                itemNum++;
                            }
                            throw new SystemException("【" + ProcessAlias + "】站,以下條碼:【" + BarcodeNo + "\n】有過站紀錄了!!!不可以啟用!!");
                        }


                        string status = "";
                        foreach (var item in resultStatus)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MoProcessChange SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MpcId = @MpcId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                MpcId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
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

        #region //UpdateMpcRoutingProcessSort -- 更新途程變更單製程排序 -- Shintokuro 2022-11-02
        public string UpdateMpcRoutingProcessSort(int MpcId, string RoutingProcessSort)
        {
            try
            {
                if (RoutingProcessSort.Length <= 0) throw new SystemException("【製程順序】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        string[] RoutingProcessSortArray = RoutingProcessSort.Split(','); //變動製程順序(全)

                        #region //撈取原途程順序
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MpcBarcodeId 
                                FROM MES.MpcBarcode a
                                WHERE a.MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);
                        var mpcBarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                        if (mpcBarcodeResult.Count() <= 0) throw new SystemException("請先在製令途程變更單綁定條碼,才能做製程順序變更!");
                        #endregion

                        #region //撈取原途程順序
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SortNumber 
                                FROM MES.MoProcess a
                                INNER JOIN MES.MoProcessChange b on a.MoId =b.MoId
                                WHERE b.MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

                        var sortNumberResult = sqlConnection.Query(sql, dynamicParameters);
                        if (sortNumberResult.Count() <= 0) throw new SystemException("途程製程資料錯誤!");
                        int itemNum = 0;
                        string sortNumbrtOre = "";
                        foreach (var item in sortNumberResult)
                        {
                            if (itemNum == 0)
                            {
                                sortNumbrtOre = Convert.ToString(item.SortNumber);
                            }
                            else
                            {
                                sortNumbrtOre += "," + Convert.ToString(item.SortNumber);
                            }
                            itemNum++;
                        }
                        string[] sortNumbrtOreArray = sortNumbrtOre.Split(','); //原製程順序
                        #endregion
                        #region //判斷順序變動的製程
                        List<string> changeSortNumbrt = new List<string>(); //變動的製程順序(比較過後)
                        itemNum = 0;
                        foreach (var item in sortNumbrtOreArray)
                        {
                            if (RoutingProcessSortArray[itemNum].Contains(item) == false)
                            {
                                changeSortNumbrt.Add(item);
                            }
                            itemNum++;

                        }
                        #endregion


                        #region //撈取變動製成的ID
                        List<int> RoutingProcessIdArray = new List<int>(); //變動的製程ID
                        bool cheack313 = true;
                        itemNum = 0;
                        foreach (var item in RoutingProcessSortArray)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.MpcId,a.MpcRoutingProcessId,a.SortNumber,b.ProcessId
                                FROM MES.MpcRoutingProcess a
                                INNER JOIN MES.MoProcess b on a.MoProcessId = b.MoProcessId
                                INNER JOIN MES.Process c on b.ProcessId = c.ProcessId
                                WHERE a.SortNumber = @SortNumber
                                AND a.MpcId = @MpcId";
                            dynamicParameters.Add("SortNumber", item);
                            dynamicParameters.Add("MpcId", MpcId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("途程製程資料錯誤!");
                            var MpcRoutingProcessId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MpcRoutingProcessId;
                            var ProcessId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ProcessId;
                            if (itemNum == 0) { if (ProcessId == 313) throw new SystemException("退鎳站不可以是首站!"); }
                            //if (ProcessId == 312) { cheack313 = true; };
                            //if (ProcessId == 313) { cheack313 = false; };
                            RoutingProcessIdArray.Add(MpcRoutingProcessId);
                            itemNum++;
                        }
                        //if (cheack313 == false) throw new SystemException("退鎳站必須在鎳站前!");
                        #endregion


                        #region //判斷途程製程資料是否正確
                        int MoProcessId = 0;
                        foreach (var item in changeSortNumbrt)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 MpcId,MoProcessId
                                FROM MES.MpcRoutingProcess
                                WHERE SortNumber = @SortNumber
                                AND MpcId = @MpcId";
                            dynamicParameters.Add("SortNumber", item);
                            dynamicParameters.Add("MpcId", MpcId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("途程製程資料錯誤!");

                            foreach (var item2 in result)
                            {
                                MoProcessId = Convert.ToInt32(item2.MoProcessId);
                            }

                            #region //判斷製成是否已經過站
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT d.BarcodeNo,a.ProcessAlias
                                    FROM MES.MoProcess a
                                    INNER JOIN MES.BarcodeProcess b ON a.MoProcessId = b.MoProcessId
	                                INNER JOIN MES.Barcode c ON b.BarcodeId = c.BarcodeId
	                                INNER JOIN MES.MpcBarcode d ON c.BarcodeNo = d.BarcodeNo
                                    WHERE a.MoProcessId = @MoProcessId";
                            dynamicParameters.Add("MoProcessId", MoProcessId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                var BarcodeNo = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).BarcodeNo;
                                var ProcessAlias = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ProcessAlias;
                                throw new SystemException("【" + BarcodeNo + "】在【" + ProcessAlias + "】已經過站，無法更動製程順序!");
                            }
                            #endregion
                        }
                        #endregion

                        #region //先將此途程製程排序改為-1
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MpcRoutingProcess SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新順序
                        int sort = 1;
                        foreach (var item3 in RoutingProcessIdArray)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.MpcRoutingProcess SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE MpcRoutingProcessId = @MpcRoutingProcessId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = sort,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    MpcRoutingProcessId = Convert.ToInt32(item3)
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            sort++;
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
        #endregion

        #region //Delete
        #region //DeleteRoutingProcess -- 刪除途程製程資料 -- Ann 2022-07-15
        public string DeleteRoutingProcess(int RoutingProcessId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷途程製程資料是否正確
                        sql = @"SELECT TOP 1 RoutingId
                                FROM MES.RoutingProcess
                                WHERE RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程製程資料錯誤!");

                        int RoutingId = 0;
                        foreach (var item in result)
                        {
                            RoutingId = Convert.ToInt32(item.RoutingId);
                        }
                        #endregion

                        #region //判斷途程是否已被確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE RoutingConfirm = @RoutingConfirm
                                AND RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingConfirm", "Y");
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("此途程已被確認，無法更動製程資料!");
                        #endregion

                        int rowsAffected = 0;
                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MoProcess
                                SET RoutingProcessId = null
                                WHERE RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a 
                                FROM MES.RoutingItemQcItem a 
                                INNER JOIN MES.RoutingItemProcess b ON a.ItemProcessId = b.ItemProcessId
                                WHERE b.RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.RoutingItemProcess
                                WHERE RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.RoutingProcess
                                WHERE RoutingProcessId = @RoutingProcessId";
                        dynamicParameters.Add("RoutingProcessId", RoutingProcessId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //重新排序
                        dynamicParameters = new DynamicParameters();
                        sql = @"With UpdateSort As
                                (
                                    SELECT SortNumber,
                                    ROW_NUMBER() OVER(ORDER BY SortNumber) NewRoutingProcessSort
                                    FROM MES.RoutingProcess
                                    WHERE RoutingId = @RoutingId
                                )
                                UPDATE MES.RoutingProcess 
                                SET SortNumber = NewRoutingProcessSort
                                FROM MES.RoutingProcess
                                INNER JOIN UpdateSort ON MES.RoutingProcess.SortNumber = UpdateSort.SortNumber
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

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

        #region //DeleteRoutingItem -- 刪除途程品號資料 -- Ann 2022-07-19
        public string DeleteRoutingItem(int RoutingItemId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷途程品號資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.RoutingItem
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程品號資料錯誤!");
                        #endregion

                        #region //判斷此途程品號是否已被製令綁定
                        sql = @"SELECT TOP 1 1
                                FROM MES.MoRouting a
                                WHERE a.RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("此途程品號已被製令綁定，無法刪除!");
                        #endregion

                        int rowsAffected = 0;
                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.RoutingItemQcItem
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.RoutingItemAttribute
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.RoutingItemProcess
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.MoRouting
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.RoutingItem
                                WHERE RoutingItemId = @RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);

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

        #region //DeleteRoutingProcessItem -- 刪除途程製程屬性檔 -- Ann 2022-08-12
        public string DeleteRoutingProcessItem(int RoutingProcessItemId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷途程製程屬性資料是否正確
                        sql = @"SELECT TOP 1 b.RoutingId
                                FROM MES.RoutingProcessItem a
                                INNER JOIN MES.RoutingProcess b ON a.RoutingProcessId = b.RoutingProcessId
                                WHERE RoutingProcessItemId = @RoutingProcessItemId";
                        dynamicParameters.Add("RoutingProcessItemId", RoutingProcessItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程製程屬性資料錯誤!");

                        int RoutingId = 0;
                        foreach (var item in result)
                        {
                            RoutingId = item.RoutingId;
                        }
                        #endregion

                        #region //判斷途程是否已被確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Routing
                                WHERE RoutingConfirm = @RoutingConfirm
                                AND RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingConfirm", "Y");
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("此途程已被確認，無法更動製程資料!");
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.RoutingProcessItem
                                WHERE RoutingProcessItemId = @RoutingProcessItemId";
                        dynamicParameters.Add("RoutingProcessItemId", RoutingProcessItemId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteRipQcItem -- 刪除途程製程加工屬性資料 -- Ann 2022-08-12
        public string DeleteRipQcItem(int ItemProcessId, int QcItemId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷途程製程屬性資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.RipQcItem a
                                WHERE ItemProcessId = @ItemProcessId
                                AND QcItemId = @QcItemId";
                        dynamicParameters.Add("ItemProcessId", ItemProcessId);
                        dynamicParameters.Add("QcItemId", QcItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程製程屬性資料錯誤!");
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.RipQcItem
                                WHERE ItemProcessId = @ItemProcessId
                                AND QcItemId = @QcItemId";
                        dynamicParameters.Add("ItemProcessId", ItemProcessId);
                        dynamicParameters.Add("QcItemId", QcItemId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteRouting -- 刪除途程資料 -- Ann 2022-11-15
        public string DeleteRouting(int RoutingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷途程資料是否正確
                        sql = @"SELECT a.RoutingConfirm
                                FROM MES.Routing a
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("途程資料錯誤!");

                        foreach (var item in result)
                        {
                            if (item.RoutingConfirm != "N") throw new SystemException("途程已被確認，無法刪除!");
                        }
                        #endregion

                        int rowsAffected = 0;
                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM MES.RoutingProcessItem a
                                INNER JOIN MES.RoutingProcess b ON a.RoutingProcessId = b.RoutingProcessId
                                WHERE b.RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM MES.RoutingProcess
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM MES.RoutingItem
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.Routing
                                WHERE RoutingId = @RoutingId";
                        dynamicParameters.Add("RoutingId", RoutingId);

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

        #region //DeleteMoProcessChange -- 刪除製令途程變更單 -- Shintokuro 2022-10-27
        public string DeleteMoProcessChange(int MpcId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //判斷製令途程變更單資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.MoProcessChange
                                WHERE MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製令途程變更單資料錯誤!");
                        #endregion


                        #region //先刪除關聯Table
                        //製令途程變更單 - 單身(變更途程製程)
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM MES.MpcRoutingProcess
                                WHERE MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        //製令途程變更單 - 單身(變更途程所綁定條碼)
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM MES.MpcBarcode
                                WHERE MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        //製令途程變更單 - 單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.MoProcessChange
                                WHERE MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

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

        #region //DeleteMpcBarcode -- 刪除製令途程變更單條碼 -- Shintokuro 2022-10-27
        public string DeleteMpcBarcode(int MpcBarcodeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //判斷製令途程變更單資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.MpcBarcode
                                WHERE MpcBarcodeId = @MpcBarcodeId";
                        dynamicParameters.Add("MpcBarcodeId", MpcBarcodeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製令途程變更單綁定條碼資料錯誤!");
                        #endregion


                        #region //先刪除關聯Table
                        //製令途程變更單 - 單身(變更途程所綁定條碼)
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM MES.MpcBarcode
                                WHERE MpcBarcodeId = @MpcBarcodeId";
                        dynamicParameters.Add("MpcBarcodeId", MpcBarcodeId);

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

        #region //DeleteMpcBarcodeAll -- 刪除製令途程變更單條碼(全) -- Shintokuro 2022-10-27
        public string DeleteMpcBarcodeAll(int MpcId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //判斷製令途程變更單資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.MpcBarcode
                                WHERE MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製令途程變更單綁定條碼資料錯誤!");
                        #endregion


                        #region //先刪除關聯Table
                        //製令途程變更單 - 單身(變更途程所綁定條碼)
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM MES.MpcBarcode
                                WHERE MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

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

        #region //DeleteMpcRoutingProcess -- 刪除製令途程變更單綁定製程 -- Shintokuro 2022-11-02
        public string DeleteMpcRoutingProcess(int MpcRoutingProcessId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //判斷製令途程變更單資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.MpcRoutingProcess
                                WHERE MpcRoutingProcessId = @MpcRoutingProcessId";
                        dynamicParameters.Add("MpcRoutingProcessId", MpcRoutingProcessId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製令途程變更單資料錯誤!");
                        #endregion


                        #region //刪除SQL - 主要Table
                        //製令途程變更單 - 單身(變更途程製程)
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM MES.MpcRoutingProcess
                                WHERE MpcRoutingProcessId = @MpcRoutingProcessId";
                        dynamicParameters.Add("MpcRoutingProcessId", MpcRoutingProcessId);

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

        #region //DeleteMpcRoutingProcessAll -- 刪除製令途程變更單綁定製程(全) -- Shintokuro 2022-11-02
        public string DeleteMpcRoutingProcessAll(int MpcId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //判斷製令途程變更單資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.MoProcessChange
                                WHERE MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製令途程變更單資料錯誤!");
                        #endregion


                        #region //刪除SQL - 主要Table
                        //製令途程變更單 - 單身(變更途程製程)
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM MES.MpcRoutingProcess
                                WHERE MpcId = @MpcId";
                        dynamicParameters.Add("MpcId", MpcId);

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
    }
}
